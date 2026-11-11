"""
Gerador do arquivo PowerPoint (.pptx) Widescreen (16:9) com Design Dark Mode
e Roteiros de Fala embutidos nas Notas do Apresentador para os 7 Integrantes.
"""

from pptx import Presentation
from pptx.util import Inches, Pt
from pptx.dml.color import RGBColor
from pptx.enum.text import PP_ALIGN
from pptx.enum.shapes import MSO_SHAPE

# Cores do Tema Dark
BG_DARK = RGBColor(11, 15, 25)        # #0b0f19
CARD_BG = RGBColor(19, 27, 46)        # #131b2e
TEXT_WHITE = RGBColor(241, 245, 249)  # #f1f5f9
TEXT_MUTED = RGBColor(148, 163, 184)  # #94a3b8
ACCENT_BLUE = RGBColor(56, 189, 248)  # #38bdf8
ACCENT_ORANGE = RGBColor(249, 115, 22)# #f97316
ACCENT_GREEN = RGBColor(16, 185, 129) # #10b981
ACCENT_AMBER = RGBColor(245, 158, 11) # #f59e0b

SLIDES_DATA = [
    {
        "meta": "SLIDE 1 • INTRODUÇÃO E CONTEXTUALIZAÇÃO",
        "title": "Apache CouchDB vs. Couchbase Server",
        "subtitle": "Arquiteturas Documentais Distribuídas: Do Offline-First à Performance Sub-Milissegundo em Memória",
        "speaker": "Integrante 1 (Tempo: 2 min | Acumulado: 02:00)",
        "col1_title": "A Crise do RDBMS Clássico",
        "col1_items": [
            "Limites de escalabilidade vertical (Scale-Up) na Web 2.0 e Mobile.",
            "Impedância Objeto-Relacional: custo de JOINs em dezenas de tabelas.",
            "Bloqueios pessimistas inviáveis em clusters distribuídos globalmente."
        ],
        "col2_title": "Dois Bancos, Duas Filosofias",
        "col2_items": [
            "Unidade básica comum: Documentos JSON auto-contidos.",
            "Apache CouchDB: Foco em HTTP/REST nativo, resiliência e sincronização Offline-First.",
            "Couchbase: Foco em arquitetura Memory-First, vBuckets e SQL++ (N1QL) em baixa latência."
        ],
        "notes": (
            "INTEGRAÇÃO DE FALA - INTEGRANTE 1 (2 minutos):\n\n"
            "Bom dia/Boa tarde a todos. Hoje nossa equipe vai apresentar uma análise arquitetural profunda sobre dois bancos de dados NoSQL orientados a documentos que frequentemente causam confusão pelo nome, mas que resolveram problemas completamente distintos na engenharia de software: o Apache CouchDB e o Couchbase.\n\n"
            "Quando olhamos para as limitações dos bancos relacionais tradicionais em sistemas distribuídos de larga escala, esbarramos em dois gargalos clássicos: a impedância objeto-relacional e o custo de manter transações distribuídas com bloqueio pessimista. Tanto o CouchDB quanto o Couchbase adotam o JSON como unidade básica de armazenamento, mas como veremos nos próximos 40 minutos, enquanto um nasceu para dominar a web descentralizada e redes instáveis, o outro foi projetado para entregar latência sub-milissegundo em clusters massivos orientados a memória."
        )
    },
    {
        "meta": "SLIDE 2 • GENEALOGIA HISTÓRICA",
        "title": "Origem Histórica: A Fusão de 2011 e a Grande Divergência",
        "subtitle": "Como CouchOne + Membase deram origem ao Couchbase enquanto o Apache CouchDB seguiu independente",
        "speaker": "Integrante 1 (Tempo: 2 min | Acumulado: 04:00)",
        "col1_title": "Linha do Tempo Histórica",
        "col1_items": [
            "2005: Damien Katz (ex-IBM Lotus Notes) cria o CouchDB (Cluster Of Unreliable Commodity Hardware).",
            "2008: CouchDB torna-se projeto Top-Level da Apache Software Foundation.",
            "2009–2010: Surgem a CouchOne (CouchDB) e a Membase (criadores do Memcached).",
            "Fevereiro 2011: Fusão CouchOne + Membase = Couchbase Inc."
        ],
        "col2_title": "A Divergência de Motores Internos",
        "col2_items": [
            "Couchbase Server NÃO é o Apache CouchDB renomeado.",
            "Couchbase uniu o motor KV em RAM do Membase com indexação JSON e SQL++.",
            "Apache CouchDB manteve seu núcleo Erlang/OTP sob governança da Apache Foundation."
        ],
        "notes": (
            "INTEGRAÇÃO DE FALA - INTEGRANTE 1 (2 minutos):\n\n"
            "Para entender por que esses dois bancos têm 'Couch' no nome, precisamos voltar a 2005. Damien Katz, que trabalhou no IBM Lotus Notes, criou o CouchDB inspirado em bancos de documentos replicados. O próprio nome COUCH é um acrônimo para Cluster Of Unreliable Commodity Hardware.\n\n"
            "Em paralelo, a equipe por trás do famoso sistema de cache distribuído Memcached criou o Membase, que adicionava persistência em disco e replicação ao Memcached. Em 2011, a empresa comercial por trás do CouchDB, chamada CouchOne, fundiu-se com a Membase, formando a Couchbase Inc.\n\n"
            "O ponto crucial para qualquer arquiteto de dados é entender o que aconteceu depois: o produto Couchbase Server não é um CouchDB com outro nome. A engenharia do Couchbase fundiu o motor de chave-valor em memória de altíssima velocidade do Membase com a capacidade de indexar JSON. Enquanto isso, a comunidade aberta manteve o Apache CouchDB evoluindo de forma independente sob a Apache Foundation."
        )
    },
    {
        "meta": "SLIDE 3 • BIFURCAÇÃO FILOSÓFICA",
        "title": "Bifurcação Filosófica: Web Nativa vs. Performance Memory-First",
        "subtitle": "Decisões opostas em protocolo de rede, controle de concorrência e caminho de I/O",
        "speaker": "Integrante 1 (Tempo: 2 min | Acumulado: 06:00)",
        "col1_title": "Apache CouchDB (Sync & Web-First)",
        "col1_items": [
            "Protocolo primário: HTTP/REST nativo (porta 5984) sem SDK proprietário obrigatório.",
            "Caminho de escrita: Disk-First (Append-Only B+Tree).",
            "Concorrência: MVCC explícito no nível do documento via árvore _rev.",
            "Foco: Replicação Master-Master, Offline-First e tolerância a partições (AP)."
        ],
        "col2_title": "Couchbase Server (Memory-First & Scale)",
        "col2_items": [
            "Protocolo primário: Binário sobre TCP (porta 11210) com SDK ciente da topologia.",
            "Caminho de escrita: Memory-First (Cache gerenciado em RAM integrado).",
            "Concorrência: vBucket único ativo (CP), CAS otimista e transações ACID.",
            "Foco: Latência sub-milissegundo, escala elástica e consultas SQL++ (N1QL)."
        ],
        "notes": (
            "INTEGRAÇÃO DE FALA - INTEGRANTE 1 (2 minutos):\n\n"
            "Essa separação gerou duas filosofias de engenharia opostas.\n"
            "De um lado, o Apache CouchDB fala HTTP de forma nativa. Você conversa com o banco diretamente usando cURL ou o navegador. Ele não usa bloqueios (locks); em vez disso, obriga a aplicação a lidar com revisões explícitas através de Multi-Version Concurrency Control (MVCC), porque seu objetivo maior é permitir que dois nós desconectados aceitem escritas e sincronizem depois sem corromper dados.\n\n"
            "Do outro lado, o Couchbase opera sobre um protocolo binário enxuto via TCP. Ele coloca uma camada de memória RAM gerenciada na frente do disco. Toda escrita vai para a RAM e retorna sucesso em microssegundos, sendo persistida de forma assíncrona em background. Agora passo a palavra para o Integrante 2, que vai dissecar o motor interno do Apache CouchDB."
        )
    },
    {
        "meta": "SLIDE 4 • ARQUITETURA INTERNA APACHE COUCHDB",
        "title": "CouchDB: Motor Append-Only e Indexação B+Tree",
        "subtitle": "Imutabilidade em disco, concorrência lock-free em Erlang e imunidade a corrupção",
        "speaker": "Integrante 2 (Tempo: 2 min | Acumulado: 08:00)",
        "col1_title": "Motor Append-Only (Tail-Append)",
        "col1_items": [
            "Escrito em Erlang/OTP: modelo de atores, alta concorrência e isolamento de falhas.",
            "Arquivos de dados (.couch) jamais sofrem sobrescrita no lugar (Zero In-Place Update).",
            "Novas revisões, nós modificados da B+Tree e o novo Database Header (Tail Header) vão para o final do arquivo."
        ],
        "col2_title": "Garantias e Processo de Compaction",
        "col2_items": [
            "Leituras nunca bloqueiam escritas (leitores acessam snapshots imutáveis da árvore B+).",
            "Imunidade a quedas de energia: basta ler de trás para frente até achar o último Tail Header íntegro.",
            "Exige Compaction periódico em background para limpar revisões obsoletas e recuperar espaço em disco."
        ],
        "notes": (
            "INTEGRAÇÃO DE FALA - INTEGRANTE 2 (2 minutos):\n\n"
            "Obrigado. Vamos mergulhar na engenharia interna do Apache CouchDB. Primeiro, ele é escrito em Erlang, a linguagem criada pela Ericsson para comutadores telefônicos, famosa por concorrência massiva e isolamento de falhas.\n\n"
            "O coração do armazenamento do CouchDB é sua estrutura Append-Only B+Tree. Diferente de bancos que sobrescrevem blocos de disco no lugar, o CouchDB jamais altera um byte já gravado. Quando você atualiza um documento ou atualiza os nós da árvore B+ de índices, ele grava os novos blocos de dados e o novo cabeçalho do banco sempre no final do arquivo.\n\n"
            "Qual é a vantagem arquitetural disso? Primeiro, leituras nunca bloqueiam escritas e escritas nunca bloqueiam leituras. Segundo, o banco é virtualmente imune à corrupção por queda brusca de energia: quando o processo sobe novamente, ele apenas varre o final do arquivo até achar o último Tail Header íntegro. O trade-off é o consumo de espaço em disco, que exige um processo de Compactação (Compaction) para reescrever o arquivo mantendo apenas a revisão mais recente."
        )
    },
    {
        "meta": "SLIDE 5 • CONTROLE DE CONCORRÊNCIA MVCC NO COUCHDB",
        "title": "CouchDB: MVCC Explícito (_rev) e Árvore de Revisões",
        "subtitle": "Como cada documento opera como um repositório Git para lidar com conflitos distribuídos",
        "speaker": "Integrante 2 (Tempo: 2 min | Acumulado: 10:00)",
        "col1_title": "Ciclo Optimistic Locking (_id + _rev)",
        "col1_items": [
            "Todo documento possui _id e _rev no formato <N>-<hash_md5> (ex: 2-7051cbe...).",
            "Para atualizar (PUT) ou remover (DELETE), o cliente é obrigado a informar o _rev atual.",
            "Se o _rev divergir da ponta da árvore no servidor, retorna imediatamente HTTP 409 Conflict.",
            "Deleções gravam uma revisão Tombstone (\"_deleted\": true) no final do arquivo."
        ],
        "col2_title": "Conflitos Determinísticos na Replicação",
        "col2_items": [
            "Edições concorrentes em nós desconectados criam ramos paralelos na árvore do documento (_conflicts).",
            "Algoritmo determinístico (baseado na geração e hash ASCII, sem depender de relógio NTP) elege o mesmo vencedor em todos os nós.",
            "Revisões perdedoras são preservadas até a aplicação realizar o merge."
        ],
        "notes": (
            "INTEGRAÇÃO DE FALA - INTEGRANTE 2 (2 minutos):\n\n"
            "Como o CouchDB garante consistência ACID no nível do documento sem usar locks pessimistas? Através de MVCC (Multi-Version Concurrency Control) exposto diretamente na API HTTP via o campo _rev.\n\n"
            "Cada documento no CouchDB não é apenas um registro, mas sim uma árvore de revisões parecida com um repositório Git. Sempre que você cria um documento, o CouchDB gera um _rev começando com 1- seguido de um hash determinístico do conteúdo. Para atualizar esse documento via PUT, você é obrigado a informar esse _rev. Se outro cliente atualizou o documento um milissegundo antes de você, o CouchDB rejeita sua escrita imediatamente com um erro HTTP 409 Conflict.\n\n"
            "Já em cenários de replicação entre dois servidores desconectados, ambos podem aceitar atualizações para o mesmo documento. Quando a rede volta e eles sincronizam, o CouchDB guarda ambas as ramificações na árvore do documento e aplica um algoritmo determinístico baseado no hash das revisões para eleger exatamente o mesmo vencedor em todos os nós do cluster."
        )
    },
    {
        "meta": "SLIDE 6 • CONSULTAS E REPLICAÇÃO NO COUCHDB",
        "title": "CouchDB: Views MapReduce, Mango (_find) e _changes Feed",
        "subtitle": "Indexação incremental em B-Tree e o protocolo de replicação contínua",
        "speaker": "Integrante 2 (Tempo: 2 min | Acumulado: 12:00)",
        "col1_title": "Mecanismos de Consulta (Views & Mango)",
        "col1_items": [
            "Design Documents (_design/): código de Views armazenado como documento no banco.",
            "MapReduce: função map(doc) emite pares emit(key, val) indexados incrementalmente; reduce calcula agregações nos nós internos da B-Tree.",
            "Mango Queries (/_find): consultas declarativas em JSON (seletores $eq, $gt, $in) introduzidas no CouchDB 2.0+."
        ],
        "col2_title": "O Endpoint /_changes e Replicação",
        "col2_items": [
            "/_changes: expõe o log sequencial de mutações do banco em tempo real (longpoll / continuous / eventsource).",
            "Couch Replication Protocol: lê o feed _changes da origem -> calcula diferenças via _revs_diff -> transfere apenas os deltas via _bulk_docs."
        ],
        "notes": (
            "INTEGRAÇÃO DE FALA - INTEGRANTE 2 (2 minutos):\n\n"
            "Para consultar dados além da chave primária _id, o CouchDB historicamente utiliza Design Documents contendo funções MapReduce escritas em JavaScript. A função map inspeciona cada documento e emite chaves e valores que são gravados em uma B-tree separada. As funções de reduce são armazenadas diretamente nos nós intermediários da B-tree. A partir da versão 2.0, o CouchDB também ganhou o Mango Query (_find), permitindo buscas declarativas em JSON.\n\n"
            "Por fim, temos a joia da coroa do CouchDB: o endpoint _changes. Como o armazenamento é append-only, o banco mantém um número de sequência incremental para cada alteração. É exatamente sobre o _changes que funciona o Protocolo de Replicação do CouchDB: um replicador lê o feed de mudanças da origem, pergunta ao destino quais revisões ele ainda não tem via _revs_diff, e envia apenas os documentos necessários. Agora passo a palavra ao Integrante 3."
        )
    },
    {
        "meta": "SLIDE 7 • ARQUITETURA INTERNA COUCHBASE SERVER",
        "title": "Couchbase: Arquitetura Memory-First e Particionamento por vBuckets",
        "subtitle": "Cache gerenciado integrado, protocolo DCP e roteamento Single-Hop sem proxies",
        "speaker": "Integrante 3 (Tempo: 2 min | Acumulado: 14:00)",
        "col1_title": "Caminho de Dados Memory-First (KV)",
        "col1_items": [
            "Leituras e escritas ocorrem diretamente no cache gerenciado em RAM (latência em microssegundos).",
            "Protocolo interno DCP (Database Change Protocol) propaga mutações em memória para:",
            "1. Replication Queue (RAM de outros nós do cluster).",
            "2. Persistence Queue (Storage engines Couchstore / Magma em disco)."
        ],
        "col2_title": "Particionamento Estático: 1024 vBuckets",
        "col2_items": [
            "Cada Bucket é dividido fixamente em 1024 vBuckets (partições lógicas).",
            "Cálculo no SDK: vBucket_ID = CRC32(Document_Key) % 1024.",
            "O SDK mantém o Cluster Map em memória e fala direto via TCP com o nó dono do vBucket ativo (Single-Hop, zero proxies)."
        ],
        "notes": (
            "INTEGRAÇÃO DE FALA - INTEGRANTE 3 (2 minutos):\n\n"
            "Obrigado. Entrando agora na engenharia do Couchbase, mudamos completamente de paradigma. Enquanto o CouchDB vai ao disco a cada operação, o Couchbase adota uma arquitetura Memory-First.\n\n"
            "Quando a sua aplicação faz um GET ou UPSERT no Couchbase, a operação acontece diretamente na memória RAM gerenciada pelo Data Service. Assim que o dado é gravado na RAM, o servidor já pode confirmar a escrita em frações de milissegundo, enquanto o protocolo interno DCP coloca essa mutação em duas filas assíncronas: uma para replicar na RAM de outros nós e outra para persistir em disco.\n\n"
            "Como o Couchbase distribui dados sem usar nós roteadores que adicionam latência? Através dos vBuckets. Todo bucket é dividido de forma fixa em 1024 partições lógicas chamadas vBuckets. O SDK aplica um hash CRC32 na chave do documento, faz módulo 1024 e consulta o Cluster Map, abrindo conexão TCP direto com o servidor onde aquele vBucket está ativo."
        )
    },
    {
        "meta": "SLIDE 8 • ESCALABILIDADE MULTI-DIMENSIONAL (MDS)",
        "title": "Couchbase: Arquitetura Multi-Service e Multi-Dimensional Scaling",
        "subtitle": "Isolamento de workloads em hardware específico para eliminar o efeito 'Noisy Neighbor'",
        "speaker": "Integrante 3 (Tempo: 2 min | Acumulado: 16:00)",
        "col1_title": "O Problema dos Clusters Homogêneos",
        "col1_items": [
            "Em bancos NoSQL tradicionais, todos os nós executam armazenamento, indexação e queries simultaneamente.",
            "Uma consulta analítica pesada consome CPU/I/O e degrada a latência das leituras operacionais por chave."
        ],
        "col2_title": "Os 6 Serviços Desacoplados do Couchbase (MDS)",
        "col2_items": [
            "1. Data Service (KV): Otimizado para alta memória RAM.",
            "2. Index Service (GSI): Otimizado para RAM e SSDs NVMe.",
            "3. Query Service (N1QL/SQL++): Otimizado para múltiplos núcleos de CPU.",
            "4. Search (FTS/Vector), 5. Analytics (MPP Colunar) e 6. Eventing (Triggers JS)."
        ],
        "notes": (
            "INTEGRAÇÃO DE FALA - INTEGRANTE 3 (2 minutos):\n\n"
            "Outro diferencial arquitetural gigantesco do Couchbase é o Multi-Dimensional Scaling (MDS), sustentado pela arquitetura Multi-Service.\n\n"
            "Na maioria dos bancos NoSQL distribuídos, cada nó do cluster executa todas as tarefas: guarda dados, atualiza índices e processa consultas pesadas. Isso cria o efeito do 'vizinho barulhento': uma query analítica complexa consome toda a CPU e derruba a latência das leituras por chave primária.\n\n"
            "O Couchbase separa o banco em seis serviços independentes: Data, Index, Query, Full-Text Search, Analytics e Eventing. Com o MDS, podemos colocar o Data Service em máquinas com muita RAM, o Query Service em máquinas com muita CPU e o Index Service em máquinas com NVMe ultrarrápidos."
        )
    },
    {
        "meta": "SLIDE 9 • N1QL (SQL++) E ÍNDICES GSI",
        "title": "Couchbase: Scopes/Collections, N1QL (SQL++) e Índices GSI",
        "subtitle": "Consultando documentos JSON aninhados com toda a expressividade do ANSI SQL",
        "speaker": "Integrante 3 (Tempo: 2 min | Acumulado: 18:00)",
        "col1_title": "Hierarquia Lógica e SQL++ (N1QL)",
        "col1_items": [
            "Hierarquia em 3 níveis (v7+): Bucket (Database) -> Scope (Schema) -> Collection (Tabela).",
            "N1QL / SQL++: Suporta SELECT, ANSI JOINs entre documentos, GROUP BY, CTEs (WITH) e transações ACID multi-documento.",
            "Operadores nativos para arrays JSON: UNNEST (achata arrays em linhas) e NEST."
        ],
        "col2_title": "GSI (Global Secondary Indexes)",
        "col2_items": [
            "Evita o padrão Scatter-Gather: índices ficam centralizados ou particionados no Index Service.",
            "Suporta Memory-Optimized Indexes (MOI) baseados em Skip-Lists lock-free em RAM.",
            "Covering Indexes permitem responder queries SQL++ diretamente do Index Service sem tocar no Data Service."
        ],
        "notes": (
            "INTEGRAÇÃO DE FALA - INTEGRANTE 3 (2 minutos):\n\n"
            "E como nós consultamos e organizamos dados no Couchbase? A partir da versão 7, o Couchbase introduziu uma hierarquia de três níveis: Bucket, Scope e Collection, permitindo mapeamento direto de bancos relacionais.\n\n"
            "Mas a grande inovação para desenvolvedores é o N1QL (pronuncia-se Nickel), hoje padronizado como SQL++. É uma implementação completa de ANSI SQL adaptada para JSON. Você escreve SELECT, WHERE, GROUP BY e até INNER JOIN entre coleções diferentes, além de operadores específicos para JSON como UNNEST, que transforma elementos de um array interno em linhas relacionais.\n\n"
            "Para que o N1QL seja rápido e não faça scatter-gather batendo em todos os nós de dados, o Couchbase usa GSI (Global Secondary Indexes). Passo agora a palavra ao Integrante 4."
        )
    },
    {
        "meta": "SLIDE 10 • POSICIONAMENTO DE MERCADO: APACHE COUCHDB",
        "title": "Onde o Apache CouchDB Brilha: O Rei do Offline-First e Edge",
        "subtitle": "Sincronização bidirecional resiliente com PouchDB e baixo consumo de hardware",
        "speaker": "Integrante 4 (Tempo: 2m45s | Acumulado: 20:45)",
        "col1_title": "O Ecossistema CouchDB + PouchDB",
        "col1_items": [
            "PouchDB implementa o protocolo CouchDB em JavaScript dentro do navegador (IndexedDB) ou mobile (SQLite).",
            "Aplicações leem e gravam localmente com latência zero mesmo 100% desconectadas da internet.",
            "Sincronização bidirecional automática em 1 linha: db.sync(remoteCouchDB, { live: true, retry: true })."
        ],
        "col2_title": "Cenários Ideais de Aplicação",
        "col2_items": [
            "Sistemas de campo sem conectividade confiável (aviação, zonas rurais, navios, plataformas e hospitais móveis).",
            "Dispositivos IoT e Gateways de Edge Computing com pouca RAM (64MB a 256MB).",
            "Catálogos e registros replicados em topologias Multi-Primary geograficamente dispersas."
        ],
        "notes": (
            "INTEGRAÇÃO DE FALA - INTEGRANTE 4 (2 minutos e 45 segundos):\n\n"
            "Obrigado. Agora que compreendemos as duas arquiteturas, surge a pergunta essencial de um arquiteto de soluções: quando escolher cada um? Vamos começar pelo Apache CouchDB.\n\n"
            "O CouchDB é imbatível quando o requisito principal do sistema é Offline-First e Sincronização na Borda (Edge). Graças ao seu protocolo de replicação sobre HTTP e ao modelo de revisões _rev, surgiu o PouchDB: um banco em JavaScript que roda direto no navegador do usuário ou dentro de um app React Native/Electron.\n\n"
            "Imagine um aplicativo usado por médicos em regiões remotas da Amazônia ou técnicos de manutenção dentro de minas subterrâneas. O aplicativo grava dados localmente no PouchDB com zero latência de rede. Quando o dispositivo encontra sinal por 30 segundos, uma única linha de código aciona a replicação contínua com o Apache CouchDB central."
        )
    },
    {
        "meta": "SLIDE 11 • CASES REAIS E TRADE-OFFS DO COUCHDB",
        "title": "Apache CouchDB no Mundo Real e Seus Trade-offs Arquiteturais",
        "subtitle": "De Blockchain corporativo (Hyperledger Fabric) a sistemas humanitários offline",
        "speaker": "Integrante 4 (Tempo: 2m45s | Acumulado: 23:30)",
        "col1_title": "Casos Reais de Mercado",
        "col1_items": [
            "Hyperledger Fabric (Linux Foundation): State Database oficial para Smart Contracts em JSON auditável.",
            "npm Registry: Arquitetura histórica de distribuição e espelhamento global de pacotes Node.js.",
            "eHealth Africa / Médicos Sem Fronteiras: Coleta epidemiológica em tablets offline durante surtos de Ebola.",
            "IBM Cloudant: Serviço DBaaS global na nuvem da IBM baseado no Apache CouchDB."
        ],
        "col2_title": "Trade-offs (Quando NÃO usar CouchDB)",
        "col2_items": [
            "Inadequado para contadores ou telemetria que atualizam o mesmo documento centenas de vezes por segundo (infla a árvore _rev e o disco).",
            "Necessidade constante de monitorar e executar Compaction em disco.",
            "Falta de JOINs relacionais e transações ACID multi-documento."
        ],
        "notes": (
            "INTEGRAÇÃO DE FALA - INTEGRANTE 4 (2 minutos e 45 segundos):\n\n"
            "Olhando para casos reais no mercado, o exemplo mais clássico do CouchDB foi o registro oficial do npm, que durante anos usou o protocolo de replicação do CouchDB para distribuir metadados de milhões de pacotes Node.js. Outro caso de altíssimo nível técnico é o Hyperledger Fabric, o principal framework de Blockchain corporativo da Linux Foundation, que adota o Apache CouchDB como banco de estado oficial (World State). Na área humanitária, organizações na África usaram CouchDB + PouchDB durante surtos de Ebola para coletar dados em tablets totalmente offline.\n\n"
            "Porém, todo arquiteto precisa conhecer os trade-offs. Você não deve escolher o CouchDB para um sistema de telemetria de alta frequência que atualiza o mesmo documento mil vezes por segundo, pois cada atualização cria uma nova folha na árvore _rev e infla o arquivo append-only até a compactação rodar. Agora passo a palavra para o Integrante 5."
        )
    },
    {
        "meta": "SLIDE 12 • POSICIONAMENTO DE MERCADO: COUCHBASE SERVER",
        "title": "Onde o Couchbase Brilha: Escala Extrema, Sub-ms e XDCR",
        "subtitle": "Unificando Cache Distribuído em Memória + Banco Documental SQL++",
        "speaker": "Integrante 5 (Tempo: 2m45s | Acumulado: 26:15)",
        "col1_title": "Eliminando a Arquitetura Dupla (Cache + DB)",
        "col1_items": [
            "Substitui a combinação complexa de RDBMS/MongoDB + Redis/Memcached por uma única plataforma.",
            "Elimina problemas de invalidação de cache e sincronização de duas camadas.",
            "Ideal para: Gerenciamento de Sessão global, Perfis 360º, Carrinhos de E-commerce, Antifraude em tempo real e Gaming."
        ],
        "col2_title": "XDCR (Cross Data Center Replication)",
        "col2_items": [
            "Replicação assíncrona Memória-a-Memória via streams binários DCP entre datacenters e nuvens (Multi-Cloud).",
            "Topologias Active-Active bidirecionais com filtragem avançada.",
            "Resolução automática de conflitos via Sequence Number (Most Updates Win) ou Timestamp com HLC (Hybrid Logical Clocks)."
        ],
        "notes": (
            "INTEGRAÇÃO DE FALA - INTEGRANTE 5 (2 minutos e 45 segundos):\n\n"
            "Obrigado. Quando olhamos para o Couchbase, entramos no território das aplicações mission-critical de altíssima vazão e latência sub-milissegundo.\n\n"
            "Em muitas arquiteturas modernas, as equipes colocam um banco documental ou relacional no fundo e precisam adicionar um cluster de Redis ou Memcached na frente para aguentar o tráfego de leitura. O Couchbase elimina essa complexidade porque ele já é o cache em memória e o banco de dados persistente com suporte a SQL++ na mesma plataforma.\n\n"
            "Além disso, para escala global entre continentes, o Couchbase possui o XDCR (Cross Data Center Replication). Diferente da replicação via HTTP do CouchDB, o XDCR transmite mutações diretamente da memória RAM de um datacenter para a memória RAM de outro datacenter usando streams binários DCP."
        )
    },
    {
        "meta": "SLIDE 13 • CASES REAIS E TRADE-OFFS DO COUCHBASE",
        "title": "Couchbase no Mundo Real e Seus Trade-offs Arquiteturais",
        "subtitle": "Operações Tier-1 de milhões de queries/segundo e custos de infraestrutura",
        "speaker": "Integrante 5 (Tempo: 2m45s | Acumulado: 29:00)",
        "col1_title": "Gigantes em Produção com Couchbase",
        "col1_items": [
            "LinkedIn: Mais de 10 milhões de queries por segundo com latência < 2ms substituindo camadas de cache + Oracle.",
            "Amadeus, Sabre & United Airlines: Inventário global de voos, assentos e check-in de passageiros em tempo real.",
            "PayPal & Revolut: Análise de fraude e sessões financeiras críticas.",
            "Domino's Pizza & Tesco: Pedidos e carrinhos em picos massivos (Super Bowl / Black Friday)."
        ],
        "col2_title": "Trade-offs (Quando NÃO usar Couchbase)",
        "col2_items": [
            "Alto requisito de memória RAM e hardware (inviável para microsserviços pequenos em VMs de 512MB/1GB).",
            "Complexidade operacional: exige planejamento de cotas de serviços MDS, índices GSI e rebalanceamento de vBuckets.",
            "Custo de licenciamento da versão Enterprise/Capella."
        ],
        "notes": (
            "INTEGRAÇÃO DE FALA - INTEGRANTE 5 (2 minutos e 45 segundos):\n\n"
            "No mercado corporativo, os números do Couchbase impressionam. O LinkedIn utiliza o Couchbase como camada crítica de dados atendendo mais de 10 milhões de operações por segundo com latência inferior a 2 milissegundos. Na indústria de aviação, gigantes como Amadeus, Sabre e United Airlines processam disponibilidade de voos e check-ins de milhões de passageiros no Couchbase. No setor financeiro, o PayPal usa Couchbase para prevenção a fraudes em tempo real.\n\n"
            "E quais são os trade-offs do Couchbase? O primeiro é o custo de hardware e memória. O Couchbase foi desenhado para servidores com dezenas ou centenas de gigabytes de RAM. O segundo é a complexidade operacional: planejar o dimensionamento de serviços Data, Index e Query exige uma equipe madura. Passo agora a palavra ao Integrante 6."
        )
    },
    {
        "meta": "SLIDE 14 • MATRIZ COMPARATIVA DIRETA (PARTE 1/2)",
        "title": "Quadro Comparativo: Modelo de Dados, CAP e Consultas",
        "subtitle": "Confronto arquitetural direto entre Apache CouchDB e Couchbase Server",
        "speaker": "Integrante 6 (Tempo: 2m45s | Acumulado: 31:45)",
        "col1_title": "Apache CouchDB",
        "col1_items": [
            "Modelo de Dados: Documentos JSON + _attachments em Databases planos.",
            "Teorema CAP / PACELC: AP (PC/EC) — prioriza disponibilidade e consistência eventual.",
            "Concorrência/ACID: MVCC explícito via _rev no documento; sem transações multi-doc.",
            "Consultas: Views MapReduce (JS/Erlang) e Mango Queries (_find). Sem JOINs."
        ],
        "col2_title": "Couchbase Server",
        "col2_items": [
            "Modelo de Dados: JSON + KV Binário em Hierarquia Bucket -> Scope -> Collection.",
            "Teorema CAP / PACELC: CP por padrão (PC/EC) — único vBucket ativo garante consistência imediata.",
            "Concorrência/ACID: CAS otimista + Locks + Transações ACID Distribuídas Multi-Documento.",
            "Consultas: N1QL / SQL++ (ANSI SQL com JOIN, UNNEST, GROUP BY) + FTS e Vector Search."
        ],
        "notes": (
            "INTEGRAÇÃO DE FALA - INTEGRANTE 6 (2 minutos e 45 segundos):\n\n"
            "Obrigado. Para sintetizar nossa análise técnica, construímos uma matriz comparativa direta dividida em seis pilares fundamentais.\n\n"
            "Primeiro, no Modelo de Dados: ambos guardam JSON, mas o Couchbase organiza os dados em Bucket, Scope e Collection, enquanto o CouchDB trabalha com bancos planos.\n"
            "Segundo, sob o Teorema CAP: o Apache CouchDB é essencialmente AP. Diante de uma partição de rede, qualquer nó aceita escritas e mantém a disponibilidade, resolvendo divergências depois via árvore de revisões. Já o Couchbase é CP por padrão: como cada chave pertence a apenas um único vBucket ativo no cluster, todas as leituras e escritas daquela chave passam pelo mesmo nó primário em memória.\n"
            "Terceiro, em Consultas: o CouchDB usa Views MapReduce ou Mango (_find) sem JOIN, enquanto o Couchbase entrega SQL++ completo com ANSI JOINs."
        )
    },
    {
        "meta": "SLIDE 15 • MATRIZ COMPARATIVA DIRETA (PARTE 2/2)",
        "title": "Quadro Comparativo: Replicação, Latência e Operação",
        "subtitle": "Critérios objetivos de decisão para Arquitetos de Software e Engenheiros de Dados",
        "speaker": "Integrante 6 (Tempo: 2m45s | Acumulado: 34:30)",
        "col1_title": "Apache CouchDB",
        "col1_items": [
            "Replicação: Couch Replication Protocol sobre HTTP (_changes + _revs_diff). Nativo para P2P e browsers (PouchDB).",
            "Performance/Latência: Casa dos milissegundos (5ms–30ms+), limitada por I/O de disco append-only e HTTP.",
            "Complexidade Operacional: Baixa infraestrutura; requer apenas gestão de Compaction e conflitos _rev."
        ],
        "col2_title": "Couchbase Server",
        "col2_items": [
            "Replicação: DCP (intra-cluster) e XDCR (inter-cluster) via TCP binário direto da RAM. Mobile via Sync Gateway + Couchbase Lite.",
            "Performance/Latência: Sub-milissegundo (microssegundos) em RAM para KV e poucos ms para N1QL com GSI.",
            "Complexidade Operacional: Moderada/Alta; exige dimensionamento de RAM, topologia MDS e rebalance de vBuckets."
        ],
        "notes": (
            "INTEGRAÇÃO DE FALA - INTEGRANTE 6 (2 minutos e 45 segundos):\n\n"
            "Continuando nossa matriz comparativa nos três últimos pilares:\n"
            "Em Mecanismos de Replicação, o protocolo HTTP do CouchDB permite que até um navegador web participe da replicação via PouchDB. O Couchbase usa o protocolo binário DCP/XDCR de altíssima velocidade entre servidores e exige o Sync Gateway e o Couchbase Lite para mobile.\n\n"
            "Em Performance e Latência: o CouchDB opera na casa dos milissegundos, limitado pelo protocolo HTTP e escrita append-only em disco; o Couchbase opera na casa de microssegundos graças ao cache gerenciado em memória.\n"
            "Na Complexidade Operacional, o CouchDB é leve e simples de subir, enquanto o Couchbase é uma plataforma enterprise que demanda planejamento de capacidade de RAM. Passo agora a palavra ao Integrante 7 para o nosso Workshop Prático."
        )
    },
    {
        "meta": "SLIDE 16 • MÓDULO 2: WORKSHOP PRÁTICO REPRODUZÍVEL",
        "title": "Arquitetura do Laboratório Prático (Docker Compose)",
        "subtitle": "Orquestrando Apache CouchDB 3.3 e Couchbase Server 7.6 no mesmo ambiente",
        "speaker": "Integrante 7 (Tempo: 2 min | Acumulado: 36:30)",
        "col1_title": "Container 1: Apache CouchDB 3.3",
        "col1_items": [
            "Imagem: apache/couchdb:3.3 (container: lab_couchdb).",
            "Porta única exposta: 5984 (API REST + Interface Web Fauxton em /_utils).",
            "Testes práticos: Criação de banco via PUT, erro HTTP 409 Conflict (_rev), Mango Query (_find) e stream contínuo /_changes."
        ],
        "col2_title": "Container 2: Couchbase Server 7.6",
        "col2_items": [
            "Imagem: couchbase:community-7.6.0 (container: lab_couchbase).",
            "Portas expostas: 8091 (Web Console), 8093 (Query N1QL), 11210 (Data KV Binário).",
            "Testes práticos: Criação de Bucket/Scope/Collection, índices GSI, queries SQL++ com UNNEST e inspeção de latência sub-ms."
        ],
        "notes": (
            "INTEGRAÇÃO DE FALA - INTEGRANTE 7 (2 minutos):\n\n"
            "Obrigado. Agora vamos sair da teoria e provar esses conceitos arquiteturais na prática. Preparamos um laboratório 100% reproduzível via Docker Compose que sobe simultaneamente o Apache CouchDB 3.3 e o Couchbase Server 7.6.\n\n"
            "Reparem na diferença imediata já na configuração de rede dos containers: para o CouchDB, precisamos expor apenas uma única porta, a 5984, porque absolutamente tudo — desde administração, inserção de dados, queries Mango até a interface web Fauxton — acontece sobre HTTP REST padrão.\n\n"
            "Já no container do Couchbase, expomos a porta 8091 para o Console de Gerenciamento, a 8093 para o motor de consultas N1QL e a 11210 para o protocolo binário de chave-valor onde os SDKs se conectam diretamente aos vBuckets."
        )
    },
    {
        "meta": "SLIDE 17 • LIVE DEMO 1: APACHE COUCHDB",
        "title": "Prática 1: Apache CouchDB via API REST, MVCC e _changes",
        "subtitle": "Executando operações nativas via cURL e monitorando eventos em tempo real",
        "speaker": "Integrante 7 (Tempo: 2m15s | Acumulado: 38:45)",
        "col1_title": "Passo a Passo REST & Conflito MVCC",
        "col1_items": [
            "1. Criar banco: curl -X PUT http://admin:admin123@localhost:5984/universidade_db",
            "2. Inserir doc: PUT /universidade_db/aluno_01 -> retorna \"rev\": \"1-4c61...\"",
            "3. Atualizar SEM _rev -> retorna HTTP 409 Conflict (Document update conflict).",
            "4. Atualizar COM _rev: \"1-4c61...\" -> avança para revisão \"2-...\"."
        ],
        "col2_title": "Mango Query (_find) e Stream /_changes",
        "col2_items": [
            "5. Criar índice JSON via POST /universidade_db/_index nos campos [\"cra\", \"curso\"].",
            "6. Consultar via POST /universidade_db/_find com seletor {\"cra\": {\"$gt\": 8.5}}.",
            "7. Monitorar mutações ao vivo no Terminal 2: GET /universidade_db/_changes?feed=continuous&since=now."
        ],
        "notes": (
            "INTEGRAÇÃO DE FALA - INTEGRANTE 7 (2 minutos e 15 segundos):\n\n"
            "Na primeira parte da nossa demonstração prática com o Apache CouchDB, vamos usar apenas o cURL no terminal para mostrar que o banco é um servidor web nativo.\n\n"
            "Primeiro criamos o banco universidade_db e inserimos o documento aluno_01. O CouchDB nos devolve rev: 1-.... Em seguida, tentamos atualizar a nota desse aluno sem passar o _rev — e vemos ao vivo o erro 409 Document update conflict, comprovando na prática o bloqueio otimista MVCC que explicamos no Slide 5. Quando passamos o _rev correto, a revisão avança para 2-....\n\n"
            "Também deixamos um segundo terminal aberto conectado ao endpoint _changes?feed=continuous. No exato instante em que alteramos o documento no primeiro terminal, o evento JSON salta na tela do segundo terminal em tempo real."
        )
    },
    {
        "meta": "SLIDE 18 • LIVE DEMO 2: COUCHBASE N1QL (SQL++) & SCRIPT PYTHON",
        "title": "Prática 2: Couchbase SQL++ (N1QL), GSI e Automação CRUD",
        "subtitle": "Consultas avançadas sobre JSON e integração programática profissional",
        "speaker": "Integrante 7 (Tempo: 2m15s | Acumulado: 41:00)",
        "col1_title": "Consultas SQL++ (N1QL) no Couchbase",
        "col1_items": [
            "Estrutura: Bucket `academico` -> Scope `graduacao` -> Collection `matriculas`.",
            "Criação de índice GSI: CREATE INDEX idx_curso_cra ON matriculas(curso, cra DESC).",
            "Consulta com UNNEST achatando array interno de disciplinas + GROUP BY por curso.",
            "Verificação no Dashboard: 0% Cache Miss Ratio e tempo de resposta sub-ms."
        ],
        "col2_title": "Script Python CRUD (crud_couchdb_demo.py)",
        "col2_items": [
            "Script automatizado executando o ciclo completo: Create -> Read -> Simulação e Tratamento de Conflito HTTP 409 -> Update com _rev -> Mango Query -> Delete (Tombstone).",
            "Demonstra como aplicações reais encapsulam o controle de revisão otimista."
        ],
        "notes": (
            "INTEGRAÇÃO DE FALA - INTEGRANTE 7 (2 minutos e 15 segundos):\n\n"
            "Na segunda etapa do nosso laboratório, entramos no Web Console do Couchbase na porta 8091. Criamos o bucket academico, o escopo graduacao e a coleção matriculas.\n\n"
            "No Query Workbench, executamos comandos SQL++ (N1QL) idênticos ao SQL que já conhecemos do mundo relacional, criando um índice GSI secundário e rodando um SELECT com filtros e UNNEST sobre arrays internos de disciplinas do aluno. Ao olharmos para o painel de estatísticas em tempo real do Couchbase, vemos que 100% das leituras vieram da memória RAM com tempo de resposta na faixa de microssegundos.\n\n"
            "Por fim, rodamos nosso script Python conectando direto no banco para demonstrar um fluxo CRUD completo com tratamento de concorrência. Com isso, concluímos nossa apresentação e abrimos para perguntas!"
        )
    }
]


def add_styled_slide(prs, data):
    blank_layout = prs.slide_layouts[6]
    slide = prs.slides.add_slide(blank_layout)

    # Background escuro
    bg = slide.background
    fill = bg.fill
    fill.solid()
    fill.fore_color.rgb = BG_DARK

    # Barra superior sutil
    top_bar = slide.shapes.add_shape(
        MSO_SHAPE.RECTANGLE, Inches(0), Inches(0), Inches(13.333), Inches(0.55)
    )
    top_bar.fill.solid()
    top_bar.fill.fore_color.rgb = RGBColor(15, 23, 42)
    top_bar.line.fill.background()

    # Texto da Meta (Esquerda da barra)
    tf_meta = top_bar.text_frame
    tf_meta.word_wrap = True
    p_meta = tf_meta.paragraphs[0]
    p_meta.text = f"  {data['meta']}   |   🎙 {data['speaker']}"
    p_meta.font.size = Pt(13)
    p_meta.font.bold = True
    p_meta.font.color.rgb = ACCENT_BLUE

    # Título do Slide (Calibrado para TV de 50 polegadas)
    title_box = slide.shapes.add_textbox(Inches(0.35), Inches(0.65), Inches(12.6), Inches(0.85))
    tf_title = title_box.text_frame
    tf_title.word_wrap = True
    p_title = tf_title.paragraphs[0]
    p_title.text = data["title"]
    p_title.font.size = Pt(31)
    p_title.font.bold = True
    p_title.font.color.rgb = TEXT_WHITE

    # Subtítulo
    sub_box = slide.shapes.add_textbox(Inches(0.35), Inches(1.38), Inches(12.6), Inches(0.5))
    tf_sub = sub_box.text_frame
    tf_sub.word_wrap = True
    p_sub = tf_sub.paragraphs[0]
    p_sub.text = data["subtitle"]
    p_sub.font.size = Pt(17.5)
    p_sub.font.color.rgb = TEXT_MUTED

    # Card Coluna 1 (Esquerda - Ocupando mais área útil da TV 50")
    card1 = slide.shapes.add_shape(
        MSO_SHAPE.ROUNDED_RECTANGLE, Inches(0.35), Inches(2.05), Inches(6.18), Inches(5.15)
    )
    card1.fill.solid()
    card1.fill.fore_color.rgb = CARD_BG
    card1.line.color.rgb = ACCENT_ORANGE
    card1.line.width = Pt(2.2)

    tf1 = card1.text_frame
    tf1.word_wrap = True
    tf1.margin_left = Inches(0.25)
    tf1.margin_right = Inches(0.25)
    tf1.margin_top = Inches(0.22)

    p1_head = tf1.paragraphs[0]
    p1_head.text = data["col1_title"]
    p1_head.font.size = Pt(22)
    p1_head.font.bold = True
    p1_head.font.color.rgb = ACCENT_ORANGE
    p1_head.space_after = Pt(14)

    for item in data["col1_items"]:
        p = tf1.add_paragraph()
        p.text = f"▸  {item}"
        p.font.size = Pt(17.5)
        p.font.color.rgb = TEXT_WHITE
        p.space_after = Pt(11)

    # Card Coluna 2 (Direita)
    card2 = slide.shapes.add_shape(
        MSO_SHAPE.ROUNDED_RECTANGLE, Inches(6.80), Inches(2.05), Inches(6.18), Inches(5.15)
    )
    card2.fill.solid()
    card2.fill.fore_color.rgb = CARD_BG
    card2.line.color.rgb = ACCENT_BLUE
    card2.line.width = Pt(2.2)

    tf2 = card2.text_frame
    tf2.word_wrap = True
    tf2.margin_left = Inches(0.25)
    tf2.margin_right = Inches(0.25)
    tf2.margin_top = Inches(0.22)

    p2_head = tf2.paragraphs[0]
    p2_head.text = data["col2_title"]
    p2_head.font.size = Pt(22)
    p2_head.font.bold = True
    p2_head.font.color.rgb = ACCENT_BLUE
    p2_head.space_after = Pt(14)

    for item in data["col2_items"]:
        p = tf2.add_paragraph()
        p.text = f"▸  {item}"
        p.font.size = Pt(17.5)
        p.font.color.rgb = TEXT_WHITE
        p.space_after = Pt(11)

    # Inserir Roteiro de Fala nas Notas do Apresentador do PowerPoint
    notes_slide = slide.notes_slide
    notes_tf = notes_slide.notes_text_frame
    notes_tf.text = data["notes"]


def main():
    prs = Presentation()
    prs.slide_width = Inches(13.333)
    prs.slide_height = Inches(7.5)

    for slide_data in SLIDES_DATA:
        add_styled_slide(prs, slide_data)

    output_path = r"d:\Área de Trabalho\BDChrouch\Apresentacao_CouchDB_vs_Couchbase.pptx"
    prs.save(output_path)
    print(f"Apresentacao PowerPoint gerada com sucesso em: {output_path}")


if __name__ == "__main__":
    main()
