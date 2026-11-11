/**
 * ============================================================================
 * MÓDULO DE SIMULADORES INTERATIVOS E AVALIAÇÃO (assets/js/simulators.js)
 * ============================================================================
 * Responsável pelas simulações visuais ao vivo nos slides de arquitetura,
 * terminais interativos do workshop e validação do Quiz universitário.
 */

(() => {
  'use strict';

  let appendRevisionCount = 2;
  const REVISION_HASHES = ['9f2', 'e18', '7b4', '3d9', 'a05'];

  /**
   * Slide 4: Simula uma escrita Append-Only no arquivo .couch do Apache CouchDB
   */
  window.simAppendCouchDB = function simAppendCouchDB() {
    appendRevisionCount += 1;
    const hash = REVISION_HASHES[appendRevisionCount % REVISION_HASHES.length];
    const strip = document.getElementById('couchdbDiskStrip');
    if (!strip) return;

    strip.innerHTML = `
      <div class="flow-node stale-block">[Doc A: rev 1-a8f]</div>
      <span class="flow-arrow">➔</span>
      <div class="flow-node stale-block">[Doc A: rev ${appendRevisionCount - 1}-c41 (Órfão)]</div>
      <span class="flow-arrow">➔</span>
      <div class="flow-node highlight-orange">[Novo Bloco: Doc A rev ${appendRevisionCount}-${hash}]</div>
      <span class="flow-arrow">➔</span>
      <div class="flow-node highlight-orange">[Nós B+Tree v${appendRevisionCount}]</div>
      <span class="flow-arrow">➔</span>
      <div class="flow-node highlight-green pulse-active"><strong>[Novo Tail Header v${appendRevisionCount} (EOF)]</strong></div>
    `;
  };

  /**
   * Slide 4: Simula o processo de Compaction limpando revisões órfãs em disco
   */
  window.simCompactCouchDB = function simCompactCouchDB() {
    const strip = document.getElementById('couchdbDiskStrip');
    if (!strip) return;

    strip.innerHTML = `
      <div class="flow-node highlight-green">✨ <strong>[Arquivo .compact Limpo]</strong></div>
      <span class="flow-arrow">➔</span>
      <div class="flow-node highlight-orange">[Apenas Revisão Ativa: Doc A rev ${appendRevisionCount}]</div>
      <span class="flow-arrow">➔</span>
      <div class="flow-node highlight-green pulse-active"><strong>[Tail Header v${appendRevisionCount} (Espaço Recuperado!)]</strong></div>
    `;
  };

  /**
   * Slide 7: Simula o cálculo de hash CRC32 -> vBucket ID e roteamento Single-Hop no Couchbase
   */
  window.simVbucket = function simVbucket(key, vb, activeNode, replicaNode) {
    const keyNode = document.getElementById('vbKeyNode');
    const activeEl = document.getElementById('vbActiveNode');
    const replicaEl = document.getElementById('vbReplicaNode');
    if (!keyNode || !activeEl || !replicaEl) return;

    keyNode.innerHTML = `SDK: <code class="inline">CRC32("${key}") % 1024 = vB #${vb}</code>`;
    activeEl.innerHTML = `<strong>RAM ${activeNode} (vB #${vb} Ativo)</strong><br/>Ack: <strong>0.28 ms</strong>`;
    replicaEl.innerHTML = `RAM ${replicaNode} (vB #${vb} Réplica) + NVMe`;
  };

  /**
   * Slides 17 & 18: Alterna a exibição do output simulado nos terminais do Workshop
   */
  window.toggleSimOutput = function toggleSimOutput(id) {
    const el = document.getElementById(id);
    if (el) el.classList.toggle('visible');
  };

  /**
   * Slides 19 & 20: Valida alternativa clicada no Quiz e revela justificativa técnica
   */
  window.checkQuiz = function checkQuiz(element, isCorrect, expId) {
    const parent = element.parentElement;
    if (parent) {
      Array.from(parent.children).forEach((opt) => {
        opt.classList.remove('correct', 'wrong');
      });
    }
    element.classList.add(isCorrect ? 'correct' : 'wrong');
    const expEl = document.getElementById(expId);
    if (expEl) expEl.classList.add('visible');
  };
})();
