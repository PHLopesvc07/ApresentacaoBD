/**
 * ============================================================================
 * MOTOR PRINCIPAL DA APRESENTAÇÃO (assets/js/app.js)
 * ============================================================================
 * Otimizado para 60 FPS em hardware de sala de aula e hospedagem estática gratuita:
 * - Cache de referências DOM (Zero queries repetidas por transição)
 * - Apontador Laser acelerado por GPU via requestAnimationFrame + translate3d
 * - Persistência leve em sessionStorage (mantém slide e escala TV em refresh)
 */

(() => {
  'use strict';

  const DEFAULT_TV_SCALE = 1.12;
  const STORAGE_KEY_SLIDE = 'nosql_deck_slide_idx';
  const STORAGE_KEY_SCALE = 'nosql_deck_tv_scale';

  // Cache único de elementos DOM
  const DOM = {
    slides: Array.from(document.querySelectorAll('.slide')),
    orb1: document.getElementById('orb1'),
    orb2: document.getElementById('orb2'),
    speakerAvatar: document.getElementById('speakerAvatar'),
    speakerLabel: document.getElementById('currentSpeakerLabel'),
    notesPanel: document.getElementById('notesPanel'),
    notesSpeaker: document.getElementById('notesSpeaker'),
    notesTime: document.getElementById('notesTime'),
    notesScript: document.getElementById('notesScript'),
    btnToggleNotes: document.getElementById('btnToggleNotes'),
    slideCounter: document.getElementById('slideCounter'),
    progressBar: document.getElementById('progressBar'),
    zoomBadge: document.getElementById('zoomBadge'),
    timerDisplay: document.getElementById('timerDisplay'),
    overviewModal: document.getElementById('overviewModal'),
    overviewGrid: document.getElementById('overviewGrid'),
    laserDot: document.getElementById('laserDot'),
    btnLaser: document.getElementById('btnLaser')
  };

  let currentIndex = parseInt(sessionStorage.getItem(STORAGE_KEY_SLIDE) || '0', 10);
  if (Number.isNaN(currentIndex) || currentIndex < 0 || currentIndex >= DOM.slides.length) {
    currentIndex = 0;
  }

  let tvScale = parseFloat(sessionStorage.getItem(STORAGE_KEY_SCALE) || String(DEFAULT_TV_SCALE));
  if (Number.isNaN(tvScale) || tvScale < 0.8 || tvScale > 1.6) {
    tvScale = DEFAULT_TV_SCALE;
  }

  // Estado do Cronômetro
  let timerSeconds = 0;
  let timerInterval = null;

  // Estado do Laser Pointer (rAF)
  let laserActive = false;
  let laserX = -100;
  let laserY = -100;
  let laserRafPending = false;

  function applyZoom() {
    document.documentElement.style.setProperty('--tv-scale', tvScale.toFixed(2));
    const pct = Math.round(tvScale * 100);
    if (DOM.zoomBadge) {
      DOM.zoomBadge.textContent = `📺 TV 50" (${pct}%)`;
    }
    try {
      sessionStorage.setItem(STORAGE_KEY_SCALE, tvScale.toFixed(2));
    } catch (_) {}
  }

  window.adjustZoom = function adjustZoom(delta) {
    tvScale = Math.min(1.60, Math.max(0.80, tvScale + delta));
    applyZoom();
  };

  window.resetZoom = function resetZoom() {
    tvScale = DEFAULT_TV_SCALE;
    applyZoom();
  };

  function updateAmbientTheme(theme) {
    if (!DOM.orb1 || !DOM.orb2) return;
    if (theme === 'couchdb') {
      DOM.orb1.style.background = 'radial-gradient(circle, #ff6b2b 0%, rgba(255, 107, 43, 0) 68%)';
      DOM.orb2.style.background = 'radial-gradient(circle, #f59e0b 0%, rgba(245, 158, 11, 0) 68%)';
    } else if (theme === 'couchbase') {
      DOM.orb1.style.background = 'radial-gradient(circle, #00d2ff 0%, rgba(0, 210, 255, 0) 68%)';
      DOM.orb2.style.background = 'radial-gradient(circle, #6366f1 0%, rgba(99, 102, 241, 0) 68%)';
    } else {
      DOM.orb1.style.background = 'radial-gradient(circle, #ff6b2b 0%, rgba(255, 107, 43, 0) 68%)';
      DOM.orb2.style.background = 'radial-gradient(circle, #00d2ff 0%, rgba(0, 210, 255, 0) 68%)';
    }
  }

  function updateSlideUI() {
    DOM.slides.forEach((slide, idx) => {
      slide.classList.toggle('active', idx === currentIndex);
    });

    const activeSlide = DOM.slides[currentIndex];
    if (!activeSlide) return;

    const speaker = activeSlide.getAttribute('data-speaker') || '';
    const timeEst = activeSlide.getAttribute('data-time') || '';
    const scriptHtml = activeSlide.getAttribute('data-script') || '';
    const theme = activeSlide.getAttribute('data-theme') || 'mixed';

    updateAmbientTheme(theme);

    const match = speaker.match(/Integrante (\d)/);
    if (DOM.speakerAvatar) DOM.speakerAvatar.textContent = match ? `I${match[1]}` : '★';
    if (DOM.speakerLabel) DOM.speakerLabel.textContent = `${speaker} • ${timeEst}`;
    if (DOM.notesSpeaker) DOM.notesSpeaker.textContent = speaker;
    if (DOM.notesTime) DOM.notesTime.textContent = `⏱ ${timeEst}`;
    if (DOM.notesScript) DOM.notesScript.innerHTML = scriptHtml;

    if (DOM.slideCounter) {
      DOM.slideCounter.textContent = `Slide ${currentIndex + 1} de ${DOM.slides.length}`;
    }
    if (DOM.progressBar) {
      const progressPct = ((currentIndex + 1) / DOM.slides.length) * 100;
      DOM.progressBar.style.width = `${progressPct}%`;
    }

    try {
      sessionStorage.setItem(STORAGE_KEY_SLIDE, String(currentIndex));
    } catch (_) {}
  }

  window.nextSlide = function nextSlide() {
    if (currentIndex < DOM.slides.length - 1) {
      currentIndex += 1;
      updateSlideUI();
    }
  };

  window.prevSlide = function prevSlide() {
    if (currentIndex > 0) {
      currentIndex -= 1;
      updateSlideUI();
    }
  };

  window.toggleNotes = function toggleNotes() {
    if (!DOM.notesPanel || !DOM.btnToggleNotes) return;
    DOM.notesPanel.classList.toggle('hidden');
    DOM.btnToggleNotes.classList.toggle('active', !DOM.notesPanel.classList.contains('hidden'));
  };

  window.toggleFullScreen = function toggleFullScreen() {
    if (!document.fullscreenElement) {
      document.documentElement.requestFullscreen().catch(() => {});
    } else {
      document.exitFullscreen();
    }
  };

  // Laser Pointer acelerado por GPU (requestAnimationFrame + translate3d)
  window.toggleLaser = function toggleLaser() {
    laserActive = !laserActive;
    document.body.classList.toggle('laser-active', laserActive);
    if (DOM.btnLaser) DOM.btnLaser.classList.toggle('active', laserActive);
  };

  function renderLaserPosition() {
    laserRafPending = false;
    if (DOM.laserDot && laserActive) {
      DOM.laserDot.style.transform = `translate3d(${laserX - 12}px, ${laserY - 12}px, 0)`;
    }
  }

  document.addEventListener('mousemove', (e) => {
    if (!laserActive) return;
    laserX = e.clientX;
    laserY = e.clientY;
    if (!laserRafPending) {
      laserRafPending = true;
      requestAnimationFrame(renderLaserPosition);
    }
  }, { passive: true });

  // Mapa Rápido de Slides (Tecla G)
  function buildOverview() {
    if (!DOM.overviewGrid) return;
    const fragment = document.createDocumentFragment();
    DOM.slides.forEach((s, idx) => {
      const title = s.querySelector('.slide-title, h1, h2')?.textContent || `Slide ${idx + 1}`;
      const spk = s.getAttribute('data-speaker') || '';
      const div = document.createElement('div');
      div.className = `overview-item ${idx === currentIndex ? 'current' : ''}`;
      div.innerHTML = `
        <div style="font-size:0.75rem; color:var(--couchbase-blue); font-weight:800; margin-bottom:4px;">SLIDE ${idx + 1} • ${spk}</div>
        <div style="font-size:0.95rem; font-weight:700; color:#fff;">${title}</div>
      `;
      div.onclick = () => {
        currentIndex = idx;
        updateSlideUI();
        window.toggleOverview();
      };
      fragment.appendChild(div);
    });
    DOM.overviewGrid.innerHTML = '';
    DOM.overviewGrid.appendChild(fragment);
  }

  window.toggleOverview = function toggleOverview() {
    if (!DOM.overviewModal) return;
    if (!DOM.overviewModal.classList.contains('open')) {
      buildOverview();
      DOM.overviewModal.classList.add('open');
    } else {
      DOM.overviewModal.classList.remove('open');
    }
  };

  // Cronômetro de Apresentação
  function startPauseTimer() {
    if (timerInterval) {
      clearInterval(timerInterval);
      timerInterval = null;
    } else {
      timerInterval = setInterval(() => {
        timerSeconds += 1;
        const mins = String(Math.floor(timerSeconds / 60)).padStart(2, '0');
        const secs = String(timerSeconds % 60).padStart(2, '0');
        if (DOM.timerDisplay) DOM.timerDisplay.textContent = `⏱ ${mins}:${secs}`;
      }, 1000);
    }
  }

  if (DOM.timerDisplay) {
    DOM.timerDisplay.addEventListener('click', startPauseTimer);
  }

  // Navegação por Teclado
  document.addEventListener('keydown', (e) => {
    if (e.key === 'Escape') {
      if (DOM.overviewModal) DOM.overviewModal.classList.remove('open');
      return;
    }
    if (e.key === 'ArrowRight' || e.key === 'PageDown' || e.key === ' ') {
      e.preventDefault();
      window.nextSlide();
    } else if (e.key === 'ArrowLeft' || e.key === 'PageUp') {
      e.preventDefault();
      window.prevSlide();
    } else if (e.key.toLowerCase() === 'r') {
      window.toggleNotes();
    } else if (e.key.toLowerCase() === 'f') {
      window.toggleFullScreen();
    } else if (e.key.toLowerCase() === 'g') {
      window.toggleOverview();
    } else if (e.key.toLowerCase() === 'l') {
      window.toggleLaser();
    } else if (e.key === '+' || e.key === '=') {
      window.adjustZoom(0.08);
    } else if (e.key === '-' || e.key === '_') {
      window.adjustZoom(-0.08);
    } else if (e.key === '0') {
      window.resetZoom();
    }
  });

  // Inicialização
  applyZoom();
  updateSlideUI();
})();
