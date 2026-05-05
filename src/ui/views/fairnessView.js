(function initializeFairnessView(globalScope) {
  function updateFairnessMethodologyCopy(options = {}) {
    const displayService = globalScope.FairnessDisplayService;
    const methodEl = options.methodEl || document.querySelector('.fairness-methodology');
    const thresholdEl = options.thresholdEl || document.querySelector('.fairness-thresholds');
    const hintEl = options.hintEl || document.getElementById('fairnessModelHint');
    const modelBadgeEls = options.modelBadgeEls || document.querySelectorAll('[data-fairness-model-badge]');

    if (!displayService || !methodEl) {
      return;
    }

    const engineType = options.engineType || globalScope.FairnessEngineService?.getSelectedFairnessEngine?.() || 'global';
    const modelLabel = displayService.getFairnessModelLabel(engineType);

    methodEl.textContent = displayService.buildFairnessMethodologyCopy(engineType);
    if (thresholdEl) {
      thresholdEl.textContent = displayService.buildFairnessThresholdCopy(engineType);
    }
    if (hintEl) {
      hintEl.textContent = engineType === 'officer_lane'
        ? 'Officer Lane Fairness evaluates fairness within officer roles and lanes.'
        : 'Global Fairness compares all officers more evenly across total opportunity.';
    }
    modelBadgeEls.forEach((el) => {
      // Shared badge update keeps the active model visible in live and simulation areas.
      el.textContent = modelLabel;
    });
  }

  function renderLiveFairnessSummaryCard(containerEl, evaluation) {
    if (!containerEl || !evaluation) {
      return;
    }

    const notes = globalScope.FairnessDisplayService?.buildFairnessNotesForDisplay(evaluation) || [];
    const modelNote = notes.find((note) => note.startsWith('Model note:')) || '';
    const detailNotes = notes.filter((note) => note !== modelNote);
    const statusLabel = String(evaluation.overallResult || 'REVIEW').toUpperCase();
    const statusClass = statusLabel.toLowerCase();
    const sections = buildFairnessSummarySections(evaluation, modelNote, detailNotes);
    const card = document.createElement('div');
    card.className = 'audit-card';
    card.innerHTML = `
      <h3>Live Fairness Summary <span class="badge badge-${escapeHtml(statusClass)}">${escapeHtml(statusLabel)}</span></h3>
      <div class="audit-summary">
        ${sections.map(renderSummarySection).join('')}
      </div>
    `;

    containerEl.appendChild(card);
  }

  function buildFairnessSummarySections(evaluation, modelNote, detailNotes) {
    const modelLabel = globalScope.FairnessDisplayService?.getFairnessModelLabel(evaluation.engineType) || 'Global Fairness';
    const sections = [];
    const overviewItems = [
      { label: 'Fairness model', value: modelLabel }
    ];
    const statusDescriptor = buildStatusDescriptorLine(evaluation?.statusMetricDescriptor);
    if (statusDescriptor) {
      overviewItems.push(statusDescriptor);
    }
    if (modelNote) {
      overviewItems.push({
        label: 'Model note',
        value: modelNote.replace(/^Model note:\s*/i, '').trim()
      });
    }
    sections.push({
      title: 'Overview',
      variant: 'overview',
      items: overviewItems
    });

    const groupedMetrics = {
      overall: [],
      consumer: [],
      mortgage: [],
      flex: [],
      policy: []
    };

    (evaluation.summaryItems || []).forEach((item) => {
      const parsed = parseSummaryLine(item);
      groupedMetrics[classifySummaryGroup(item)].push(parsed);
    });

    if (groupedMetrics.overall.length) {
      sections.push({ title: 'Overall', items: groupedMetrics.overall });
    }
    if (groupedMetrics.consumer.length) {
      sections.push({ title: 'Consumer Lane', items: groupedMetrics.consumer });
    }
    if (groupedMetrics.mortgage.length) {
      sections.push({ title: 'Mortgage Lane', items: groupedMetrics.mortgage });
    }
    if (groupedMetrics.flex.length) {
      sections.push({ title: 'Flex Lane', items: groupedMetrics.flex });
    }
    if (groupedMetrics.policy.length) {
      sections.push({ title: 'Policy And Thresholds', items: groupedMetrics.policy });
    }

    if (detailNotes.length) {
      sections.push({
        title: 'Interpretation',
        variant: 'notes',
        notes: detailNotes
      });
    }

    return sections;
  }

  function renderSummarySection(section) {
    const sectionClass = section.variant ? ` audit-summary-section-${escapeHtml(section.variant)}` : '';
    const itemsHtml = (section.items || []).map((item) => `
      <div class="audit-summary-item">
        <div class="audit-summary-item-label">${escapeHtml(item.label)}</div>
        <div class="audit-summary-item-value">${escapeHtml(item.value)}</div>
      </div>
    `).join('');
    const notesHtml = (section.notes || []).map((note) => `
      <li class="audit-summary-note">${escapeHtml(String(note))}</li>
    `).join('');
    return `
      <section class="audit-summary-section${sectionClass}">
        <div class="audit-summary-section-title">${escapeHtml(section.title)}</div>
        ${itemsHtml ? `<div class="audit-summary-grid">${itemsHtml}</div>` : ''}
        ${notesHtml ? `<ul class="audit-summary-notes">${notesHtml}</ul>` : ''}
      </section>
    `;
  }

  function buildStatusDescriptorLine(descriptor) {
    if (!descriptor || !descriptor.label) {
      return null;
    }
    const contextLabel = descriptor.contextLabel ? `${descriptor.contextLabel}: ` : '';
    const numericValue = Number(descriptor.valuePercent);
    const suffix = Number.isFinite(numericValue) ? ` ${numericValue.toFixed(1)}%` : '';
    return {
      label: 'Primary status driver',
      value: `${contextLabel}${descriptor.label}${suffix}`.trim()
    };
  }

  function parseSummaryLine(text) {
    const rawText = String(text || '').trim();
    const separatorIndex = rawText.indexOf(':');
    if (separatorIndex === -1) {
      return { label: 'Detail', value: rawText };
    }
    return {
      label: rawText.slice(0, separatorIndex).trim(),
      value: rawText.slice(separatorIndex + 1).trim()
    };
  }

  function classifySummaryGroup(text) {
    const normalized = String(text || '').trim().toLowerCase();
    if (normalized.startsWith('heloc') || normalized.includes('routing') || normalized.includes('advisory band applied')) {
      return 'policy';
    }
    if (normalized.startsWith('overall ')) {
      return 'overall';
    }
    if (normalized.startsWith('consumer ')) {
      return 'consumer';
    }
    if (normalized.startsWith('mortgage ')) {
      return 'mortgage';
    }
    if (normalized.startsWith('flex ')) {
      return 'flex';
    }
    return 'overall';
  }

  function escapeHtml(value) {
    return String(value)
      .replaceAll('&', '&amp;')
      .replaceAll('<', '&lt;')
      .replaceAll('>', '&gt;')
      .replaceAll('"', '&quot;')
      .replaceAll("'", '&#39;');
  }

  globalScope.FairnessView = {
    updateFairnessMethodologyCopy,
    renderLiveFairnessSummaryCard
  };
})(window);
