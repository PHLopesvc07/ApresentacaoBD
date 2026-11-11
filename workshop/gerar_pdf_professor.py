"""
Gerador de PDF Acadêmico e Técnico para Entrega ao Professor:
- Capa Universitária e Ficha Técnica do Projeto
- Estrutura Completa dos Slides e Roteiros de Fala dos 7 Integrantes
- Matriz Comparativa Detalhada (CAP, PACELC, MVCC, vBuckets)
- Roteiro Completo do Workshop Prático com Docker e Python
- Avaliação com 5 Questões Inéditas e Gabarito Comentado
- Plano de Contingência e Ensaio de 48 Horas
"""

import os
from reportlab.lib.pagesizes import A4
from reportlab.lib import colors
from reportlab.lib.styles import getSampleStyleSheet, ParagraphStyle
from reportlab.platypus import (
    SimpleDocTemplate, Paragraph, Spacer, Table, TableStyle, PageBreak, KeepTogether, HRFlowable
)
from reportlab.pdfgen import canvas

PDF_PATH = r"d:\Área de Trabalho\BDChrouch\Apresentacao_CouchDB_vs_Couchbase_Professor.pdf"

# Paleta Profissional para Relatório PDF
C_PRIMARY = colors.HexColor("#0f172a")      # Azul ardósia escuro
C_SECONDARY = colors.HexColor("#1e293b")    # Cinza escuro
C_ACCENT_BLUE = colors.HexColor("#0284c7")  # Ciano/Azul técnico
C_ACCENT_ORANGE = colors.HexColor("#ea580c")# Laranja CouchDB
C_ACCENT_GREEN = colors.HexColor("#059669") # Verde sucesso
C_BG_LIGHT = colors.HexColor("#f8fafc")     # Fundo de tabelas claro
C_TEXT_DARK = colors.HexColor("#0f172a")    # Texto principal
C_TEXT_MUTED = colors.HexColor("#475569")   # Texto secundário
C_BORDER = colors.HexColor("#cbd5e1")       # Bordas sutis


class NumberedCanvas(canvas.Canvas):
    """Canvas com numeração de página e cabeçalho dinâmico."""
    def __init__(self, *args, **kwargs):
        super().__init__(*args, **kwargs)
        self._saved_page_states = []

    def showPage(self):
        self._saved_page_states.append(dict(self.__dict__))
        self._startPage()

    def save(self):
        num_pages = len(self._saved_page_states)
        for state in self._saved_page_states:
            self.__dict__.update(state)
            self.draw_page_decorations(num_pages)
            super().showPage()
        super().save()

    def draw_page_decorations(self, total_pages):
        if self._pageNumber == 1:
            return  # Capa limpa sem cabeçalho/rodapé

        self.saveState()
        self.setFont("Helvetica-Bold", 8)
        self.setFillColor(C_TEXT_MUTED)

        # Cabeçalho Superior
        self.drawString(54, 800, "SEMINÁRIO ACADÊMICO: APACHE COUCHDB vs. COUCHBASE SERVER")
        self.drawRightString(541, 800, "ARQUITETURA DE BANCOS DE DADOS NOSQL")
        self.setStrokeColor(C_BORDER)
        self.setLineWidth(0.6)
        self.line(54, 794, 541, 794)

        # Rodapé Inferior
        self.line(54, 45, 541, 45)
        self.setFont("Helvetica", 8)
        self.drawString(54, 34, "Documento de Entrega Avaliativa • Ciência da Computação / Engenharia de Software")
        self.drawRightString(541, 34, f"Página {self._pageNumber} de {total_pages}")
        self.restoreState()


def build_pdf():
    doc = SimpleDocTemplate(
        PDF_PATH,
        pagesize=A4,
        leftMargin=54,
        rightMargin=54,
        topMargin=54,
        bottomMargin=54
    )

    styles = getSampleStyleSheet()

    # Estilos Customizados
    style_cover_title = ParagraphStyle(
        'CoverTitle',
        fontName='Helvetica-Bold',
        fontSize=26,
        leading=32,
        textColor=C_PRIMARY,
        alignment=0,
        spaceAfter=10
    )
    style_cover_sub = ParagraphStyle(
        'CoverSubtitle',
        fontName='Helvetica',
        fontSize=13,
        leading=18,
        textColor=C_TEXT_MUTED,
        spaceAfter=20
    )
    style_h1 = ParagraphStyle(
        'SectionH1',
        fontName='Helvetica-Bold',
        fontSize=16,
        leading=21,
        textColor=C_PRIMARY,
        spaceBefore=16,
        spaceAfter=10,
        keepWithNext=True
    )
    style_h2 = ParagraphStyle(
        'SectionH2',
        fontName='Helvetica-Bold',
        fontSize=12,
        leading=16,
        textColor=C_ACCENT_BLUE,
        spaceBefore=12,
        spaceAfter=6,
        keepWithNext=True
    )
    style_body = ParagraphStyle(
        'BodyMain',
        fontName='Helvetica',
        fontSize=9.5,
        leading=14,
        textColor=C_TEXT_DARK,
        spaceAfter=6
    )
    style_bullet = ParagraphStyle(
        'BulletMain',
        fontName='Helvetica',
        fontSize=9,
        leading=13.5,
        textColor=C_TEXT_DARK,
        leftIndent=15,
        spaceAfter=4
    )
    style_script = ParagraphStyle(
        'ScriptSpeech',
        fontName='Helvetica-Oblique',
        fontSize=9,
        leading=13.5,
        textColor=colors.HexColor("#1e3a8a"),
        spaceBefore=3,
        spaceAfter=3
    )
    style_code = ParagraphStyle(
        'CodeStyle',
        fontName='Courier',
        fontSize=8,
        leading=11,
        textColor=colors.HexColor("#0f172a")
    )
    style_table_header = ParagraphStyle(
        'TableHeader',
        fontName='Helvetica-Bold',
        fontSize=9,
        leading=12,
        textColor=colors.white,
        alignment=1
    )
    style_table_cell = ParagraphStyle(
        'TableCell',
        fontName='Helvetica',
        fontSize=8.5,
        leading=11.5,
        textColor=C_TEXT_DARK
    )

    story = []

    # =========================================================================
    # CAPA DO RELATÓRIO DO PROFESSOR
    # =========================================================================
    story.append(Spacer(1, 40))
    story.append(Paragraph("UNIVERSIDADE / DEPARTAMENTO DE COMPUTAÇÃO", ParagraphStyle('Inst', fontName='Helvetica-Bold', fontSize=10, leading=14, textColor=C_TEXT_MUTED)))
    story.append(Paragraph("DISCIPLINA: BANCO DE DADOS AVANÇADO / SISTEMAS DISTRIBUÍDOS", ParagraphStyle('Disc', fontName='Helvetica-Bold', fontSize=10, leading=14, textColor=C_ACCENT_BLUE)))
    story.append(Spacer(1, 30))

    story.append(Paragraph("Seminário Acadêmico & Laboratório Prático:<br/>Apache CouchDB vs. Couchbase Server", style_cover_title))
    story.append(HRFlowable(width="100%", thickness=3, color=C_ACCENT_BLUE, spaceBefore=4, spaceAfter=14))
    story.append(Paragraph("Dossiê Técnico Completo: Fundamentos Arquiteturais, Roteiros de Fala dos 7 Integrantes, Laboratório com Docker, Matriz Comparativa e Avaliação Didática.", style_cover_sub))
    story.append(Spacer(1, 20))

    # Ficha Técnica da Apresentação
    meta_table_data = [
        [Paragraph("<b>Duração Prevista:</b>", style_table_cell), Paragraph("35 a 45 minutos no total", style_table_cell)],
        [Paragraph("<b>Equipe de Alunos:</b>", style_table_cell), Paragraph("7 Integrantes (fala equilibrada de ~5 a 6 min/integrante)", style_table_cell)],
        [Paragraph("<b>Público-Alvo:</b>", style_table_cell), Paragraph("Graduação em Ciência da Computação / Engenharia de Software", style_table_cell)],
        [Paragraph("<b>Repositório Online:</b>", style_table_cell), Paragraph("https://github.com/PHLopesvc07/ApresentacaoBD", style_table_cell)],
        [Paragraph("<b>Apresentação Web:</b>", style_table_cell), Paragraph("https://phlopesvc07.github.io/ApresentacaoBD/", style_table_cell)],
        [Paragraph("<b>Slides PowerPoint:</b>", style_table_cell), Paragraph("Apresentacao_CouchDB_vs_Couchbase.pptx (16:9 Dark Mode)", style_table_cell)],
    ]
    t_meta = Table(meta_table_data, colWidths=[130, 357])
    t_meta.setStyle(TableStyle([
        ('BACKGROUND', (0,0), (-1,-1), C_BG_LIGHT),
        ('BOX', (0,0), (-1,-1), 1, C_BORDER),
        ('INNERGRID', (0,0), (-1,-1), 0.5, C_BORDER),
        ('TOPPADDING', (0,0), (-1,-1), 6),
        ('BOTTOMPADDING', (0,0), (-1,-1), 6),
    ]))
    story.append(t_meta)
    story.append(Spacer(1, 30))

    # Sumário dos Módulos
    story.append(Paragraph("<b>Estrutura do Dossiê para Avaliação Docente:</b>", style_h2))
    story.append(Paragraph("• <b>Módulo 1:</b> Roteiro de Fala e Estrutura dos 18 Slides Técnicos (Integrantes 1 a 7)", style_bullet))
    story.append(Paragraph("• <b>Módulo 2:</b> Matriz Comparativa Completa (CAP, PACELC, MVCC, N1QL e Operação)", style_bullet))
    story.append(Paragraph("• <b>Módulo 3:</b> Guia do Workshop Prático (Docker Compose, cURL, N1QL e Script Python)", style_bullet))
    story.append(Paragraph("• <b>Módulo 4:</b> Avaliação Didática (5 Questões Universitárias Inéditas com Gabarito Comentado)", style_bullet))
    story.append(Paragraph("• <b>Módulo 5:</b> Plano Operacional de Contingência e Ensaio Cronometrado (48 Horas)", style_bullet))

    story.append(PageBreak())

    # =========================================================================
    # MÓDULO 1: DISTRIBUIÇÃO DA EQUIPE E SLIDES COM ROTEIROS DE FALA
    # =========================================================================
    story.append(Paragraph("MÓDULO 1: PESQUISA, SLIDES E ROTEIROS DE FALA", style_h1))
    story.append(Paragraph("Planejamento estrito de fala para os 7 integrantes da equipe, mantendo tempo equilibrado e postura técnica universitária.", style_body))

    # Tabela de Cronograma da Equipe
    team_table_data = [
        [Paragraph("Integrante", style_table_header), Paragraph("Slides", style_table_header), Paragraph("Tema Responsável", style_table_header), Paragraph("Tempo Est.", style_table_header)],
        [Paragraph("<b>Integrante 1</b>", style_table_cell), Paragraph("1, 2 e 3", style_table_cell), Paragraph("Origem Histórica, Genealogia (Membase + CouchOne) e Bifurcação Filosófica", style_table_cell), Paragraph("~6 min", style_table_cell)],
        [Paragraph("<b>Integrante 2</b>", style_table_cell), Paragraph("4, 5 e 6", style_table_cell), Paragraph("Arquitetura Apache CouchDB: Append-Only B+Tree, MVCC (_rev), Views e _changes", style_table_cell), Paragraph("~6 min", style_table_cell)],
        [Paragraph("<b>Integrante 3</b>", style_table_cell), Paragraph("7, 8 e 9", style_table_cell), Paragraph("Arquitetura Couchbase: Memory-First, vBuckets, MDS (Multi-Service) e N1QL/GSI", style_table_cell), Paragraph("~6 min", style_table_cell)],
        [Paragraph("<b>Integrante 4</b>", style_table_cell), Paragraph("10 e 11", style_table_cell), Paragraph("Casos de Uso e Mercado CouchDB: PouchDB Offline-First, IoT e Trade-offs", style_table_cell), Paragraph("~5m30s", style_table_cell)],
        [Paragraph("<b>Integrante 5</b>", style_table_cell), Paragraph("12 e 13", style_table_cell), Paragraph("Casos de Uso e Mercado Couchbase: Escala Extrema, XDCR, LinkedIn e Trade-offs", style_table_cell), Paragraph("~5m30s", style_table_cell)],
        [Paragraph("<b>Integrante 6</b>", style_table_cell), Paragraph("14 e 15", style_table_cell), Paragraph("Matriz Comparativa Técnica Direta: Teorema CAP/PACELC, Latência e Operação", style_table_cell), Paragraph("~5m30s", style_table_cell)],
        [Paragraph("<b>Integrante 7</b>", style_table_cell), Paragraph("16, 17 e 18", style_table_cell), Paragraph("Workshop Prático ao Vivo: Docker Compose, REST API, SQL++ e Script Python", style_table_cell), Paragraph("~6m30s", style_table_cell)],
    ]
    t_team = Table(team_table_data, colWidths=[75, 45, 307, 60])
    t_team.setStyle(TableStyle([
        ('BACKGROUND', (0,0), (-1,0), C_PRIMARY),
        ('GRID', (0,0), (-1,-1), 0.5, C_BORDER),
        ('TOPPADDING', (0,0), (-1,-1), 5),
        ('BOTTOMPADDING', (0,0), (-1,-1), 5),
        ('ROWBACKGROUNDS', (0,1), (-1,-1), [colors.white, C_BG_LIGHT]),
    ]))
    story.append(t_team)
    story.append(Spacer(1, 14))

    # Detalhamento de cada Integrante com Slides e Script
    slides_info = [
        (
            "Integrante 1 — Slides 1 a 3 (Fundamentos e Genealogia Histórica)",
            [
                ("Slide 1: Capa e Contextualização do Problema",
                 "• A crise do RDBMS clássico na Web 2.0 e limitações de escala vertical (Scale-Up).\n• Impedância objeto-relacional e o custo de bloqueios pessimistas distribuídos.\n• O modelo documental em JSON e a divergência entre dois bancos com nomes similares.",
                 "\"Bom dia/Boa tarde a todos. Hoje nossa equipe vai apresentar uma análise arquitetural profunda sobre dois bancos de dados NoSQL orientados a documentos que frequentemente causam confusão pelo nome, mas que resolveram problemas completamente distintos: o Apache CouchDB e o Couchbase. Quando olhamos para as limitações dos bancos relacionais tradicionais em sistemas distribuídos de larga escala, esbarramos em dois gargalos clássicos: a impedância objeto-relacional e o custo de manter transações com bloqueio pessimista. CouchDB e Couchbase adotam JSON, mas enquanto um nasceu para dominar redes instáveis e offline-first, o outro foi desenhado para latência sub-milissegundo em memória.\""),
                ("Slide 2: A Fusão de 2011 e a Grande Divergência",
                 "• 2005: Damien Katz cria o CouchDB (Cluster Of Unreliable Commodity Hardware) em Erlang.\n• 2010: Criação do Membase (Memcached + persistência + vBuckets).\n• Fevereiro 2011: Fusão CouchOne + Membase = Couchbase Inc.\n• Divergência: Couchbase adotou motor em RAM do Membase; Apache CouchDB manteve motor Erlang na ASF.",
                 "\"Para entender por que ambos têm 'Couch' no nome, voltamos a 2005. O CouchDB foi criado por Damien Katz com foco em hardware comum. Em paralelo, criadores do Memcached criaram o Membase. Em 2011, fundiram-se na Couchbase Inc. O ponto vital é: o produto Couchbase Server não é o CouchDB renomeado. A engenharia fundiu o motor em memória do Membase com JSON, enquanto a Apache Software Foundation manteve o CouchDB independente.\""),
                ("Slide 3: Bifurcação Filosófica Fundamental",
                 "• Apache CouchDB: Protocolo nativo HTTP/REST (porta 5984), caminho Disk-First (append-only) e MVCC explícito via _rev.\n• Couchbase: Protocolo binário TCP (porta 11210), caminho Memory-First (RAM), consistência por vBucket (CP) e N1QL (SQL++).",
                 "\"Essa separação gerou duas filosofias opostas. O Apache CouchDB fala HTTP nativamente, grava em disco append-only e expõe o campo _rev para concorrência otimista e sincronização offline. O Couchbase opera em protocolo binário TCP, colocando um cache gerenciado em RAM na frente do disco para confirmar operações em frações de milissegundo. Passo agora ao Integrante 2 para dissecar o CouchDB.\"")
            ]
        ),
        (
            "Integrante 2 — Slides 4 a 6 (Arquitetura Interna do Apache CouchDB)",
            [
                ("Slide 4: Motor Append-Only e B+Tree Storage",
                 "• Escrito em Erlang/OTP: Concorrência massiva e isolamento de falhas (Let it crash).\n• Zero in-place updates: Modificações vão para o fim do arquivo .couch junto com o Tail Header.\n• Imunidade a corrupção em falhas elétricas e necessidade periódica de Compaction.",
                 "\"O motor de armazenamento do CouchDB em Erlang é uma B+Tree Append-Only. Ele nunca sobrescreve blocos existentes; novas revisões e nós modificados vão para o final do arquivo com o novo Tail Header. Isso traz duas vantagens fundamentais: leituras nunca bloqueiam escritas, e o banco é virtualmente imune a corrupções por queda de energia. O contraponto é a necessidade de Compactação para limpar dados órfãos.\""),
                ("Slide 5: Controle MVCC (_rev) e Árvore de Revisões",
                 "• Identificador de revisão no formato <geração>-<hash_md5>.\n• Atualizações exigem _rev atual da ponta da folha; divergências disparam HTTP 409 Conflict.\n• Conflitos em replicação offline bifurcam a árvore; algoritmo determinístico elege o vencedor sem NTP.",
                 "\"O CouchDB garante consistência através de MVCC no documento. Para atualizar um registro, a aplicação deve enviar o _rev atual; caso contrário, recebe HTTP 409 Conflict. Em replicação desconectada, se dois nós editarem o mesmo registro, o CouchDB preserva ambas as revisões e usa um algoritmo determinístico baseado no hash para escolher a mesma versão vencedora em todos os nós.\""),
                ("Slide 6: Views MapReduce, Mango Query e _changes",
                 "• Views em Design Documents (_design/): Funções map(doc) em JS emitindo chaves para B-Trees.\n• Mango Queries (/_find): Sintaxe declarativa em JSON introduzida no CouchDB 2.0+.\n• Endpoint /_changes: Log sequencial de mutações que alimenta o protocolo de replicação contínua.",
                 "\"Para consultas, temos Views MapReduce e Mango Queries declarativas em JSON. O grande trunfo é o endpoint _changes: um log sequencial de alterações em tempo real. O protocolo de replicação lê esse feed, compara revisões via _revs_diff e transfere apenas as diferenças em lote. Passo agora a palavra ao Integrante 3.\"")
            ]
        ),
        (
            "Integrante 3 — Slides 7 a 9 (Arquitetura Interna do Couchbase Server)",
            [
                ("Slide 7: Arquitetura Memory-First e 1024 vBuckets",
                 "• Data Service (KV): Leituras e escritas servidas diretamente do Managed Cache em RAM.\n• Protocolo interno DCP (Database Change Protocol): Stream binário em memória para replicação e disco.\n• Particionamento fixo em 1024 vBuckets: SDK calcula CRC32(Key) % 1024 e conecta via Single-Hop.",
                 "\"No Couchbase, o caminho de dados é Memory-First. As escritas vão para a RAM gerenciada pelo Data Service e retornam confirmação imediata em microssegundos, enquanto o protocolo DCP propaga os dados para réplicas e disco. O bucket é fatiado em 1024 vBuckets estáticos; o SDK calcula o hash CRC32 e conecta diretamente no nó ativo daquele vBucket sem proxies intermediários.\""),
                ("Slide 8: Multi-Dimensional Scaling (MDS) e Isolamento",
                 "• Fim do vizinho barulhento (Noisy Neighbor) típico de nós homogêneos.\n• Desacoplamento em 6 serviços: Data (RAM), Index (NVMe), Query (CPU), Search, Analytics e Eventing.\n• Escala independente de hardware conforme o gargalo real de cada workload.",
                 "\"Em bancos NoSQL clássicos, todos os nós fazem tudo. No Couchbase, o Multi-Dimensional Scaling separa o cluster em serviços independentes: Data com muita RAM, Query com muita CPU e Index com SSDs ultrarrápidos. Mutações fluem via memória pelo protocolo DCP sem que consultas analíticas pesadas roubem CPU das leituras de chave-valor.\""),
                ("Slide 9: Scopes, Collections, N1QL (SQL++) e GSI",
                 "• Mapeamento relacional: Bucket (DB) -> Scope (Schema) -> Collection (Tabela).\n• N1QL / SQL++: ANSI SQL completo para JSON com suporte a JOIN, UNNEST, CTEs e transações ACID.\n• Global Secondary Indexes (GSI): Evita consultas lentas do tipo Scatter-Gather.",
                 "\"O Couchbase organiza dados em Bucket, Scope e Collection. Sua linguagem N1QL (SQL++) permite consultas ANSI SQL em JSON, incluindo JOINs e o operador UNNEST para desmembrar arrays. Para garantir que o SQL++ seja veloz, os índices GSI ficam centralizados nos nós de indexação, respondendo queries em milissegundos. Passo ao Integrante 4.\"")
            ]
        ),
        (
            "Integrante 4 — Slides 10 e 11 (Casos de Mercado e Trade-offs do CouchDB)",
            [
                ("Slide 10: O Rei do Offline-First (CouchDB + PouchDB)",
                 "• PouchDB no cliente: Executa no navegador ou app móvel com sincronização bidirecional em 1 linha.\n• Zero latência percebida: Aplicações operam sem rede e sincronizam deltas quando houver conexão.\n• Cenários ideais: Mineração, saúde em zonas rurais, aviação, plataformas offshore e IoT.",
                 "\"Quando escolher o CouchDB? Quando a premissa é Offline-First e borda. Pareado ao PouchDB, ele permite que aplicações web e mobile funcionem perfeitamente sem sinal de internet. Imagine equipes de saúde em aldeias remotas ou técnicos em minas: eles coletam dados offline e, ao obter sinal por 30 segundos, o banco sincroniza com a nuvem automaticamente.\""),
                ("Slide 11: Cases Reais e Limitações do CouchDB",
                 "• Casos reais: Hyperledger Fabric (World State oficial em JSON), npm registry e eHealth Africa.\n• Trade-offs: Inadequado para contadores de alta frequência na mesma chave e escritas massivas contínuas.\n• Falta de JOINs no servidor e necessidade de monitoramento de espaço em disco e compactação.",
                 "\"O CouchDB é usado pelo Hyperledger Fabric da Linux Foundation como banco oficial de estado de contratos inteligentes em JSON auditável. Porém, não use o CouchDB para telemetria de alta frequência que atualiza o mesmo documento 500 vezes por segundo, pois a árvore de revisões e o arquivo append-only inflarão com rapidez. Passo ao Integrante 5.\"")
            ]
        ),
        (
            "Integrante 5 — Slides 12 e 13 (Casos de Mercado e Trade-offs do Couchbase)",
            [
                ("Slide 12: Escala Extrema e Replicação Global XDCR",
                 "• Eliminação da arquitetura dupla: Substitui a combinação complexa RDBMS/MongoDB + Redis.\n• XDCR (Cross Data Center Replication): Stream binário memória-a-memória entre datacenters/nuvens.\n• Suporte a topologias Active-Active e resolução de conflitos via HLC (Hybrid Logical Clocks).",
                 "\"O Couchbase brilha em missões críticas de altíssima escala. Ele elimina a complexidade de manter um banco no fundo e um cluster de Redis na frente. Para escala global, o XDCR replica dados diretamente da memória RAM de um datacenter no Brasil para outro nos Estados Unidos via protocolo DCP, resolvendo conflitos em tempo real com relógios lógicos híbridos.\""),
                ("Slide 13: Cases Globais e Limitações do Couchbase",
                 "• Cases reais: LinkedIn (10 milhões de queries/s com < 2ms), Amadeus/Sabre (aviação global) e PayPal.\n• Trade-offs: Alto custo de hardware (exige dezenas de GBs de RAM) e complexidade em Day-2 Ops.\n• Inviável para microsserviços pequenos em VMs de 512 MB/1 GB.",
                 "\"Os números do Couchbase impressionam: o LinkedIn processa 10 milhões de queries por segundo, e empresas como Amadeus gerenciam reservas de voos no mundo todo. Mas o custo é alto: ele exige servidores robustos com muita RAM e planejamento refinado de serviços MDS. Passo a palavra ao Integrante 6 para a síntese comparativa.\"")
            ]
        ),
        (
            "Integrante 6 — Slides 14 e 15 (Matriz Comparativa Técnica Direta)",
            [
                ("Slide 14: Comparativo de Modelo, CAP e Consultas",
                 "• Modelo: CouchDB (Databases planos) vs. Couchbase (Bucket ➔ Scope ➔ Collection).\n• Teorema CAP: CouchDB é AP (disponibilidade contínua); Couchbase é CP por padrão (consistência por vBucket).\n• Consultas: Views/Mango sem JOIN no CouchDB vs. SQL++ com ANSI JOIN e GSI no Couchbase.",
                 "\"Sintetizando os dois bancos em nossa matriz comparativa: no modelo de dados, o Couchbase espelha o relacional em Bucket, Scope e Collection. Sob o Teorema CAP, o CouchDB prioriza disponibilidade (AP), enquanto o Couchbase foca em consistência imediata servida pelo nó ativo do vBucket (CP por padrão). Em consultas, o Couchbase entrega o poder do SQL++.\""),
                ("Slide 15: Comparativo de Replicação, Latência e Operação",
                 "• Replicação: HTTP/REST no CouchDB (PouchDB amigável) vs. TCP binário DCP/XDCR em RAM no Couchbase.\n• Latência: Casa de milissegundos no CouchDB vs. Sub-milissegundo em microssegundos no Couchbase.\n• Conclusão de decisão técnica: Conectividade instável pede CouchDB; throughput e latência crítica pedem Couchbase.",
                 "\"Nos aspectos operacionais: a replicação do CouchDB sobre HTTP permite integração simples com navegadores. O Couchbase opera em protocolo binário TCP direto da memória, atingindo latência sub-milissegundo. Em suma: se o desafio é rede e offline-first, vá de CouchDB; se o desafio é vazão e latência, vá de Couchbase. Passo ao Integrante 7 para o workshop prático.\"")
            ]
        ),
        (
            "Integrante 7 — Slides 16 a 18 (Workshop Prático e Demonstração ao Vivo)",
            [
                ("Slide 16: Topologia do Laboratório Docker Compose",
                 "• Container 1: apache/couchdb:3.3 exposto na porta 5984 (API REST e interface Fauxton /_utils).\n• Container 2: couchbase:community-7.6.0 com portas 8091 (Web Console), 8093 (Query) e 11210 (KV).",
                 "\"Agora vamos ver a arquitetura em execução no laboratório Docker Compose. O CouchDB requer apenas a porta 5984 para API e interface web Fauxton. Já o Couchbase expõe a porta 8091 para administração, 8093 para queries N1QL e 11210 para o acesso binário dos SDKs aos vBuckets.\""),
                ("Slide 17: Live Demo 1 — Apache CouchDB (MVCC e _changes)",
                 "• Demonstração de inserção via cURL e falha com HTTP 409 Conflict ao omitir o _rev.\n• Atualização com sucesso ao enviar a revisão correta e inspeção do stream contínuo /_changes.",
                 "\"Na primeira demonstração, criamos o aluno_01 no CouchDB. Ao tentar atualizar a nota sem o _rev atual, o servidor devolve na hora o erro HTTP 409 Document update conflict, provando o MVCC ao vivo. Passando o _rev correto, o dado atualiza e o evento salta no terminal que escuta o feed contínuo do endpoint _changes.\""),
                ("Slide 18: Live Demo 2 — Couchbase SQL++ e Automação Python",
                 "• Criação de índice GSI, execução de query SQL++ com UNNEST em 1.4 ms e script Python CRUD completo.\n• Verificação de 0% Cache Miss Ratio (leitura 100% da RAM) e abertura para dinâmica com a turma.",
                 "\"No Couchbase, entramos no Query Workbench, criamos o índice secundário GSI e executamos uma consulta com UNNEST sobre o array de disciplinas. A resposta retorna em 1.4 milissegundos direto da RAM. Concluímos com nosso script Python demonstrando o tratamento de concorrência em código real e abrindo para a avaliação!\"")
            ]
        )
    ]

    for member_title, member_slides in slides_info:
        story.append(Paragraph(member_title, style_h2))
        for s_title, s_bullets, s_script in member_slides:
            card_content = [
                [Paragraph(f"<b>{s_title}</b>", style_table_cell)],
                [Paragraph(s_bullets.replace('\n', '<br/>'), style_bullet)],
                [Paragraph(f"<b>Roteiro Falado (Script):</b> {s_script}", style_script)]
            ]
            t_slide = Table(card_content, colWidths=[487])
            t_slide.setStyle(TableStyle([
                ('BACKGROUND', (0,0), (-1,-1), C_BG_LIGHT),
                ('BOX', (0,0), (-1,-1), 1, C_BORDER),
                ('TOPPADDING', (0,0), (-1,-1), 5),
                ('BOTTOMPADDING', (0,0), (-1,-1), 5),
                ('LINEBELOW', (0,0), (-1,0), 0.5, C_BORDER),
                ('LINEBELOW', (0,1), (-1,1), 0.5, C_BORDER),
            ]))
            story.append(t_slide)
            story.append(Spacer(1, 8))
        story.append(Spacer(1, 6))

    story.append(PageBreak())

    # =========================================================================
    # MÓDULO 2: MATRIZ COMPARATIVA TÉCNICA
    # =========================================================================
    story.append(Paragraph("MÓDULO 2: MATRIZ COMPARATIVA TÉCNICA DIRETA", style_h1))
    story.append(Paragraph("Confronto detalhado das garantias arquiteturais sob as lentes do Teorema CAP, teorema PACELC e engenharia de dados.", style_body))

    matrix_data = [
        [Paragraph("Dimensão Técnica", style_table_header), Paragraph("Apache CouchDB", style_table_header), Paragraph("Couchbase Server", style_table_header)],
        [
            Paragraph("<b>Modelo de Dados</b>", style_table_cell),
            Paragraph("Documentos JSON + Anexos binários (_attachments). Estrutura plana de Databases sem schemas intermediários.", style_table_cell),
            Paragraph("JSON + Chave-Valor Binário. Hierarquia em 3 níveis: Bucket (DB) ➔ Scope (Schema) ➔ Collection (Tabela).", style_table_cell)
        ],
        [
            Paragraph("<b>Teorema CAP & PACELC</b>", style_table_cell),
            Paragraph("<b>AP (PC/EC):</b> Prioriza disponibilidade contínua e tolerância a partições. Consistência eventual com resolução determinística de conflitos.", style_table_cell),
            Paragraph("<b>CP por padrão (PC/EC):</b> Um único nó ativo detém a posse do vBucket por chave, garantindo consistência imediata (Read-Your-Own-Writes) em RAM.", style_table_cell)
        ],
        [
            Paragraph("<b>Concorrência & ACID</b>", style_table_cell),
            Paragraph("MVCC otimista explícito via árvore de revisões (_rev). Sem suporte a transações multi-documento no servidor.", style_table_cell),
            Paragraph("CAS (Compare-And-Swap) otimista ou bloqueio pessimista + Transações ACID Distribuídas Multi-Documento (v6.5+).", style_table_cell)
        ],
        [
            Paragraph("<b>Mecanismos de Consulta</b>", style_table_cell),
            Paragraph("Views pré-computadas em B-Tree (MapReduce em JS/Erlang) + Mango Query (POST /_find) declarativo em JSON. Não possui JOINs.", style_table_cell),
            Paragraph("<b>N1QL / SQL++:</b> ANSI SQL completo para JSON com suporte nativo a INNER/LEFT JOIN, UNNEST (arrays), GROUP BY e índices GSI.", style_table_cell)
        ],
        [
            Paragraph("<b>Mecanismos de Replicação</b>", style_table_cell),
            Paragraph("Couch Replication Protocol sobre HTTP/REST (_changes + _revs_diff). Nativo para peer-to-peer, offline e navegadores (PouchDB).", style_table_cell),
            Paragraph("Protocolo DCP (intra-cluster) e XDCR (inter-cluster) via stream binário TCP de memória para memória. Mobile requer Sync Gateway.", style_table_cell)
        ],
        [
            Paragraph("<b>Performance e Latência</b>", style_table_cell),
            Paragraph("Casa dos <b>milissegundos (5ms a 30ms+)</b>, limitada pelo I/O append-only de disco e serialização de cabeçalhos HTTP.", style_table_cell),
            Paragraph("Latência <b>sub-milissegundo (< 1ms / microssegundos)</b> servida diretamente do cache gerenciado em RAM para operações de chave.", style_table_cell)
        ],
        [
            Paragraph("<b>Complexidade Operacional</b>", style_table_cell),
            Paragraph("Baixa pegada de memória. Roda com < 128 MB de RAM. Exige rotina de Compaction para liberar espaço em disco de revisões órfãs.", style_table_cell),
            Paragraph("Exige servidores robustos (16 GB a 64 GB+ de RAM por nó), planejamento de serviços MDS e rebalanceamento de vBuckets.", style_table_cell)
        ],
    ]
    t_matrix = Table(matrix_data, colWidths=[105, 191, 191])
    t_matrix.setStyle(TableStyle([
        ('BACKGROUND', (0,0), (-1,0), C_PRIMARY),
        ('GRID', (0,0), (-1,-1), 0.5, C_BORDER),
        ('TOPPADDING', (0,0), (-1,-1), 6),
        ('BOTTOMPADDING', (0,0), (-1,-1), 6),
        ('ROWBACKGROUNDS', (0,1), (-1,-1), [colors.white, C_BG_LIGHT]),
        ('VALIGN', (0,0), (-1,-1), 'TOP'),
    ]))
    story.append(t_matrix)

    story.append(PageBreak())

    # =========================================================================
    # MÓDULO 3: ROTEIRO DO WORKSHOP PRÁTICO REPRODUZÍVEL
    # =========================================================================
    story.append(Paragraph("MÓDULO 3: WORKSHOP PRÁTICO (PASSO A PASSO REPRODUZÍVEL)", style_h1))
    story.append(Paragraph("Guia de comandos testados e funcionais utilizando Docker Compose para execução em laboratório universitário.", style_body))

    story.append(Paragraph("1. Orquestração com Docker Compose (docker-compose.yml)", style_h2))
    story.append(Paragraph("Sobe simultaneamente o Apache CouchDB 3.3 e o Couchbase Server 7.6 Community Edition:", style_body))

    compose_code = """version: '3.8'
services:
  couchdb:
    image: apache/couchdb:3.3
    container_name: lab_couchdb
    ports:
      - "5984:5984" # API HTTP REST + Interface Fauxton (/_utils)
    environment:
      - COUCHDB_USER=admin
      - COUCHDB_PASSWORD=admin123

  couchbase:
    image: couchbase:community-7.6.0
    container_name: lab_couchbase
    ports:
      - "8091-8096:8091-8096" # Web Console (8091), N1QL Query (8093)
      - "11210:11210"         # Data Service KV Binário (vBuckets)"""

    t_code1 = Table([[Paragraph(compose_code.replace('\n', '<br/>').replace(' ', '&nbsp;'), style_code)]], colWidths=[487])
    t_code1.setStyle(TableStyle([
        ('BACKGROUND', (0,0), (-1,-1), colors.HexColor("#f1f5f9")),
        ('BOX', (0,0), (-1,-1), 1, C_BORDER),
        ('TOPPADDING', (0,0), (-1,-1), 6),
        ('BOTTOMPADDING', (0,0), (-1,-1), 6),
    ]))
    story.append(t_code1)
    story.append(Spacer(1, 10))

    story.append(Paragraph("2. Comandos de Inicialização dos Clusters", style_h2))
    init_code = """# Subir os containers em background
docker compose up -d

# Inicializar Single-Node Setup no Apache CouchDB 3.3 (Criação de _users e _replicator)
curl -X POST http://admin:admin123@localhost:5984/_cluster_setup \\
  -H "Content-Type: application/json" \\
  -d '{"action": "enable_single_node", "username": "admin", "password": "admin123", "bind_address": "0.0.0.0", "port": 5984, "singlenode": true}'

# Inicializar o cluster Couchbase Server 7.6 via CLI com serviços Data, Index e Query
docker exec lab_couchbase couchbase-cli cluster-init \\
  -c 127.0.0.1 --cluster-username admin --cluster-password admin123 \\
  --services data,index,query --cluster-ramsize 1024 --cluster-index-ramsize 512 --index-storage-setting default"""

    t_code2 = Table([[Paragraph(init_code.replace('\n', '<br/>').replace(' ', '&nbsp;'), style_code)]], colWidths=[487])
    t_code2.setStyle(TableStyle([
        ('BACKGROUND', (0,0), (-1,-1), colors.HexColor("#f1f5f9")),
        ('BOX', (0,0), (-1,-1), 1, C_BORDER),
        ('TOPPADDING', (0,0), (-1,-1), 6),
        ('BOTTOMPADDING', (0,0), (-1,-1), 6),
    ]))
    story.append(t_code2)
    story.append(Spacer(1, 10))

    story.append(Paragraph("3. Comandos Práticos no Apache CouchDB (REST, MVCC e Stream)", style_h2))
    couch_cmds = """# Criar o banco de dados
curl -X PUT http://admin:admin123@localhost:5984/universidade_db

# Inserir documento com ID fixo -> Devolve "rev": "1-4c61..."
curl -X PUT http://admin:admin123@localhost:5984/universidade_db/aluno_01 \\
  -H "Content-Type: application/json" \\
  -d '{"nome": "Ana Carolina Silva", "curso": "Ciencia da Computacao", "cra": 8.7}'

# PROVOCAÇÃO DE ERRO MVCC: Atualizar sem _rev -> Dispara HTTP 409 Conflict!
curl -i -X PUT http://admin:admin123@localhost:5984/universidade_db/aluno_01 \\
  -H "Content-Type: application/json" -d '{"nome": "Ana Carolina Silva", "cra": 9.1}'

# Atualização correta enviando o _rev -> Sucesso e avanço para rev "2-7051..."
curl -X PUT http://admin:admin123@localhost:5984/universidade_db/aluno_01 \\
  -H "Content-Type: application/json" \\
  -d '{"_rev": "1-4c6114c65e295552ab1019e2b046b10e", "nome": "Ana Carolina Silva", "cra": 9.1}'

# Escutar feed de alterações contínuo em tempo real (Terminal 2)
curl -N "http://admin:admin123@localhost:5984/universidade_db/_changes?feed=continuous&include_docs=true&since=now\""""

    t_code3 = Table([[Paragraph(couch_cmds.replace('\n', '<br/>').replace(' ', '&nbsp;'), style_code)]], colWidths=[487])
    t_code3.setStyle(TableStyle([
        ('BACKGROUND', (0,0), (-1,-1), colors.HexColor("#f1f5f9")),
        ('BOX', (0,0), (-1,-1), 1, C_BORDER),
        ('TOPPADDING', (0,0), (-1,-1), 6),
        ('BOTTOMPADDING', (0,0), (-1,-1), 6),
    ]))
    story.append(t_code3)
    story.append(Spacer(1, 10))

    story.append(Paragraph("4. Comandos N1QL (SQL++) no Couchbase Server", style_h2))
    couchbase_cmds = """-- Criação do Bucket via CLI e Scopes/Collections via SQL++
CREATE SCOPE `academico`.`graduacao` IF NOT EXISTS;
CREATE COLLECTION `academico`.`graduacao`.`matriculas` IF NOT EXISTS;

-- Criação de Índices GSI (Primário e Secundário Composto)
CREATE PRIMARY INDEX `idx_primary` ON `academico`.`graduacao`.`matriculas`;
CREATE INDEX `idx_curso_cra` ON `academico`.`graduacao`.`matriculas`(`curso`, `cra` DESC);

-- Consulta Analítica com Operador UNNEST (Achatamento de Array JSON)
SELECT m.nome AS aluno, d.nome AS disciplina, d.nota
FROM `academico`.`graduacao`.`matriculas` AS m
UNNEST m.disciplinas AS d
WHERE m.curso = "Ciencia da Computacao" AND m.cra >= 9.0;"""

    t_code4 = Table([[Paragraph(couchbase_cmds.replace('\n', '<br/>').replace(' ', '&nbsp;'), style_code)]], colWidths=[487])
    t_code4.setStyle(TableStyle([
        ('BACKGROUND', (0,0), (-1,-1), colors.HexColor("#f1f5f9")),
        ('BOX', (0,0), (-1,-1), 1, C_BORDER),
        ('TOPPADDING', (0,0), (-1,-1), 6),
        ('BOTTOMPADDING', (0,0), (-1,-1), 6),
    ]))
    story.append(t_code4)

    story.append(PageBreak())

    # =========================================================================
    # MÓDULO 4: AVALIAÇÃO DIDÁTICA (5 QUESTÕES UNIVERSITÁRIAS + GABARITO)
    # =========================================================================
    story.append(Paragraph("MÓDULO 4: AVALIAÇÃO DIDÁTICA UNIVERSITÁRIA", style_h1))
    story.append(Paragraph("5 Questões inéditas de múltipla escolha com gabarito oficial e justificativas técnicas completas.", style_body))

    questions = [
        (
            "Questão 1 — Controle de Concorrência MVCC e Armazenamento no Apache CouchDB",
            "Em uma arquitetura baseada no Apache CouchDB, dois clientes concorrentes (Cliente A e Cliente B) leem simultaneamente o documento '_id': 'pedido_99', cuja revisão é '1-967a00d...'. O Cliente A submete uma requisição PUT com o _rev lido e recebe HTTP 201 Created com nova revisão '2-7051cbe...'. Um milissegundo depois, o Cliente B envia sua requisição PUT para o mesmo documento informando ainda a revisão '1-967a00d...'. O que ocorre internamente?",
            [
                ("A", "O CouchDB bloqueia o Cliente B em fila pessimista e sobrescreve no mesmo bloco de disco (in-place update)."),
                ("B", "O CouchDB rejeita imediatamente o Cliente B com HTTP 409 Conflict, mantendo a escrita do Cliente A gravada de forma append-only no final do arquivo até a execução do Compaction."),
                ("C", "O CouchDB aceita ambas as escritas aplicando Last-Write-Wins (LWW) baseado no relógio NTP do servidor."),
                ("D", "O CouchDB realiza merge automático dos campos JSON alterados se as propriedades forem distintas."),
                ("E", "O CouchDB grava a atualização na memória RAM do vBucket correspondente e agenda a resolução para o Index Service.")
            ],
            "B",
            "O CouchDB implementa MVCC estrito. Ao receber um _rev obsoleto em relação à folha ativa do documento, ele rejeita imediatamente a operação com HTTP 409 Conflict. Por usar armazenamento append-only, os bytes da revisão antiga permanecem em disco até que ocorra um processo de compactação (Compaction)."
        ),
        (
            "Questão 2 — Particionamento por vBuckets e Roteamento no Couchbase Server",
            "Uma equipe de engenharia projeta um sistema financeiro de baixa latência em um cluster Couchbase de 8 nós executando o Data Service. Sobre o mecanismo de particionamento e roteamento KV, assinale a correta:",
            [
                ("A", "O Couchbase utiliza proxies roteadores centrais semelhantes ao mongos consultando nós ZooKeeper."),
                ("B", "Cada Bucket sofre shard splitting dinâmico a cada 64 MB, exigindo que o cliente consulte a API REST a cada leitura."),
                ("C", "Cada Bucket é particionado estaticamente em exatamente 1024 vBuckets, onde o hash CRC32 da chave permite ao SDK consultar seu Cluster Map e conectar-se diretamente ao nó que detém o vBucket ativo via TCP (Single-Hop)."),
                ("D", "Os vBuckets operam em modo Multi-Primary com todos os 8 nós aceitando escritas simultâneas para a mesma chave."),
                ("E", "Quando um novo nó é adicionado, o cluster é obrigado a dobrar o total de vBuckets de 1024 para 2048.")
            ],
            "C",
            "O particionamento em 1024 vBuckets é estático. A função CRC32(chave) % 1024 determina o vBucket. Como o SDK cliente mantém o Cluster Map atualizado em sua própria memória, ele abre conexão TCP direta com o nó onde aquele vBucket está ativo, sem intermediários."
        ),
        (
            "Questão 3 — Multi-Dimensional Scaling (MDS) e Consultas SQL++ (N1QL)",
            "Qual é a principal vantagem arquitetural da abordagem Multi-Dimensional Scaling (MDS) no Couchbase Server?",
            [
                ("A", "Elimina o uso de memória RAM no Data Service, passando todo o tráfego para arquivos de log em disco."),
                ("B", "Permite isolar cargas em hardware especializado (Data em RAM, Index em NVMe e Query em CPU) e evita consultas do tipo Scatter-Gather através de índices GSI centralizados."),
                ("C", "Obriga que consultas SQL++ sejam pré-compiladas em Design Documents JavaScript como no CouchDB."),
                ("D", "Força atualizações bloqueantes e síncronas dos índices secundários no caminho crítico de escrita KV."),
                ("E", "Substitui o protocolo binário DCP por chamadas HTTP GET /_changes em texto plano entre os nós.")
            ],
            "B",
            "O MDS desacopla os serviços do banco. Nós de Data focam em memória para operações chave-valor; nós de Query focam em múltiplos núcleos de CPU para resolver joins; e nós de Index mantêm GSIs em SSDs NVMe rápidos, alimentados assincronamente pelo protocolo de streaming DCP sem degradar o tráfego KV."
        ),
        (
            "Questão 4 — Teorema CAP e Arquitetura Offline-First",
            "Considere: Cenário I (App de inspeção em pistas remotas sem internet com sync local) e Cenário II (Carrinho de compras global com 500k ops/s e consistência imediata). Relacionando ao Teorema CAP e aos dois bancos, assinale a correta:",
            [
                ("A", "O Cenário I deve usar Couchbase conectado diretamente ao browser pela porta 11210; o Cenário II deve usar CouchDB."),
                ("B", "Ambos os bancos são classificados estritamente como CA, sendo incapazes de operar diante de partições de rede."),
                ("C", "O Cenário I alinha-se ao Apache CouchDB (AP), pareado ao PouchDB via HTTP e árvore _rev; o Cenário II alinha-se ao Couchbase (CP por padrão), servindo operações direto da RAM do vBucket ativo."),
                ("D", "No Cenário I, o CouchDB descarta edições offline conflitantes sem guardar histórico de ramificação."),
                ("E", "No Cenário II, o CouchDB tem menor latência que o Couchbase por gravar diretamente em disco append-only.")
            ],
            "C",
            "O CouchDB é essencialmente AP, perfeito para Offline-First com PouchDB no cliente. O Couchbase adota modelo CP por padrão no nível do vBucket (um único primário ativo por chave), servindo requisições da memória RAM com consistência imediata e latência de microssegundos."
        ),
        (
            "Questão 5 — Deleção Lógica (Tombstones), Compactação e XDCR",
            "Sobre exclusão de dados e replicação no Apache CouchDB e no Couchbase Server, assinale a alternativa correta:",
            [
                ("A", "Deletar um documento no CouchDB apaga fisicamente os bytes de disco na hora, liberando espaço sem Compaction."),
                ("B", "No CouchDB, exclusões geram uma nova revisão Tombstone ('_deleted': true) gravada append-only para que a replicação propague a remoção para outros nós em vez de ressuscitar o registro."),
                ("C", "O mecanismo XDCR do Couchbase não suporta replicação Active-Active e trava em caso de conflitos concorrentes."),
                ("D", "A compactação do CouchDB apaga documentos que não receberam leituras nos últimos 30 dias."),
                ("E", "CouchDB e Couchbase exigem downtime total de leituras e escritas para rodar compactação de armazenamento.")
            ],
            "B",
            "Em bancos append-only replicados, deletar dados requer a gravação de um Tombstone ('_deleted': true). Sem esse registro histórico propagado pelo feed de alterações (_changes), nós réplicas desatualizados interpretariam a ausência do dado como uma falha e ressuscitariam o documento."
        ),
    ]

    for q_title, q_desc, q_opts, q_ans, q_exp in questions:
        story.append(Paragraph(f"<b>{q_title}</b>", style_h2))
        story.append(Paragraph(q_desc, style_body))
        for opt_letter, opt_text in q_opts:
            story.append(Paragraph(f"<b>({opt_letter})</b> {opt_text}", style_bullet))
        story.append(Spacer(1, 3))
        # Caixa de Gabarito
        ans_box = [
            [Paragraph(f"<b>Gabarito Oficial: Alternativa ({q_ans})</b>", style_table_cell)],
            [Paragraph(f"<b>Justificativa Técnica:</b> {q_exp}", style_script)]
        ]
        t_ans = Table(ans_box, colWidths=[487])
        t_ans.setStyle(TableStyle([
            ('BACKGROUND', (0,0), (-1,-1), colors.HexColor("#ecfdf5")),
            ('BOX', (0,0), (-1,-1), 1, C_ACCENT_GREEN),
            ('TOPPADDING', (0,0), (-1,-1), 4),
            ('BOTTOMPADDING', (0,0), (-1,-1), 4),
        ]))
        story.append(t_ans)
        story.append(Spacer(1, 10))

    story.append(PageBreak())

    # =========================================================================
    # MÓDULO 5: PLANO DE CONTINGÊNCIA (48 HORAS)
    # =========================================================================
    story.append(Paragraph("MÓDULO 5: PLANO OPERACIONAL DE CONTINGÊNCIA (48 HORAS)", style_h1))
    story.append(Paragraph("Matriz de papéis e cronograma de ensaio para garantir apresentação impecável de 35 a 45 minutos.", style_body))

    roles_data = [
        [Paragraph("Integrante", style_table_header), Paragraph("Papel de Apresentação", style_table_header), Paragraph("Função Operacional nas 48 Horas", style_table_header)],
        [Paragraph("<b>Integrante 1</b>", style_table_cell), Paragraph("Slides 1 a 3 (Origem & Filosofia)", style_table_cell), Paragraph("Líder de Tempo: Monitora cronômetro e passagens de bastão nos ensaios.", style_table_cell)],
        [Paragraph("<b>Integrante 2</b>", style_table_cell), Paragraph("Slides 4 a 6 (Motor CouchDB)", style_table_cell), Paragraph("Revisor Técnico CouchDB: Valida funcionamento de B+Tree, MVCC e _changes.", style_table_cell)],
        [Paragraph("<b>Integrante 3</b>", style_table_cell), Paragraph("Slides 7 a 9 (Motor Couchbase)", style_table_cell), Paragraph("Revisor Técnico Couchbase: Valida conceitos de vBuckets, MDS e N1QL.", style_table_cell)],
        [Paragraph("<b>Integrante 4</b>", style_table_cell), Paragraph("Slides 10 e 11 (Mercado CouchDB)", style_table_cell), Paragraph("Designer do Deck: Cuida da diagramação e formatação visual dos slides.", style_table_cell)],
        [Paragraph("<b>Integrante 5</b>", style_table_cell), Paragraph("Slides 12 e 13 (Mercado Couchbase)", style_table_cell), Paragraph("Co-Designer e Handout: Prepara impressão e distribuição do Quiz de avaliação.", style_table_cell)],
        [Paragraph("<b>Integrante 6</b>", style_table_cell), Paragraph("Slides 14 e 15 (Matriz Comparativa)", style_table_cell), Paragraph("Engenheiro de Backup: Grava vídeo de 4 min com comandos do workshop (Plano B).", style_table_cell)],
        [Paragraph("<b>Integrante 7</b>", style_table_cell), Paragraph("Slides 16 a 18 (Workshop Prático)", style_table_cell), Paragraph("Piloto de Live Demo: Testa Docker Compose, cURL, portas e script Python.", style_table_cell)],
    ]
    t_roles = Table(roles_data, colWidths=[80, 160, 247])
    t_roles.setStyle(TableStyle([
        ('BACKGROUND', (0,0), (-1,0), C_PRIMARY),
        ('GRID', (0,0), (-1,-1), 0.5, C_BORDER),
        ('TOPPADDING', (0,0), (-1,-1), 5),
        ('BOTTOMPADDING', (0,0), (-1,-1), 5),
        ('ROWBACKGROUNDS', (0,1), (-1,-1), [colors.white, C_BG_LIGHT]),
    ]))
    story.append(t_roles)
    story.append(Spacer(1, 14))

    story.append(Paragraph("Checkpoints de Cronômetro da Apresentação (Meta: 41 minutos)", style_h2))
    story.append(Paragraph("• <b>Fim do Integrante 1 (Slide 3):</b> 06:00 acumulados", style_bullet))
    story.append(Paragraph("• <b>Fim do Integrante 2 (Slide 6):</b> 12:00 acumulados", style_bullet))
    story.append(Paragraph("• <b>Fim do Integrante 3 (Slide 9):</b> 18:00 acumulados", style_bullet))
    story.append(Paragraph("• <b>Fim do Integrante 4 (Slide 11):</b> 23:30 acumulados", style_bullet))
    story.append(Paragraph("• <b>Fim do Integrante 5 (Slide 13):</b> 29:00 acumulados", style_bullet))
    story.append(Paragraph("• <b>Fim do Integrante 6 (Slide 15):</b> 34:30 acumulados", style_bullet))
    story.append(Paragraph("• <b>Fim do Integrante 7 / Workshop (Slide 18):</b> 41:00 acumulados (Margem de 4 minutos para perguntas!)", style_bullet))

    doc.build(story, canvasmaker=NumberedCanvas)
    print(f"PDF acadêmico gerado com sucesso em: {PDF_PATH}")


if __name__ == "__main__":
    build_pdf()
