import { CATEGORIES, ALL_OPERATIONS } from '../catalog/index.js';
import { escapeHtml } from '../core/security.js';
import { setState } from '../core/store.js';

export function renderHomeView(container, { navigate }) {
  const catPreview = CATEGORIES.slice(0, 6)
    .map(
      (c) => `
      <article class="category-card">
        <div class="cat-header">
          <span class="cat-code" style="border: 1px solid ${escapeHtml(c.color)}">${escapeHtml(c.code)}</span>
          <span class="badge badge-neutral">${ALL_OPERATIONS.filter((o) => o.category === c.id).length} opérations</span>
        </div>
        <h3>${escapeHtml(c.name)}</h3>
        <p class="field-hint">${escapeHtml(c.description)}</p>
        <div class="op-example">${escapeHtml(c.example)}</div>
        <div class="card-actions">
          <button type="button" class="btn-secondary btn-sm" data-open-cat="${escapeHtml(c.id)}">Explorer</button>
        </div>
      </article>
    `
    )
    .join('');

  container.innerHTML = `
    <div class="stack-lg">
      <section class="hero-banner">
        <div>
          <span class="hero-badge">100 % en français · Hors-ligne &amp; Confidentialité totale · Windows, Mac &amp; Web</span>
          <h1 class="page-title">Apprenez Excel, automatisez en VBA ou Office Scripts, en toute fiabilité</h1>
          <p class="page-subtitle">
            Environnement complet de travail et de formation professionnelle : Académie Excel progressive (avec et sans code), Assistant d'automatisation en français, Audit VBA et Analyseur de classeurs.
          </p>
          <div class="hero-actions">
            <button type="button" class="btn-primary" data-nav="formulas">Laboratoire &amp; Formules Excel (39)</button>
            <button type="button" class="btn-secondary" data-nav="tutorial">Ouvrir l'Académie Excel &amp; VBA</button>
            <button type="button" class="btn-secondary" data-nav="free">Lancer l'Assistant automatisation</button>
            <button type="button" class="btn-secondary" data-nav="guided">Mode rapide (129 opérations)</button>
            <button type="button" class="btn-secondary" data-nav="auditor">Auditer / Corriger une macro VBA</button>
            <button type="button" class="btn-ghost" data-nav="analyzer">Analyser un fichier Excel</button>
          </div>
        </div>
        <div class="hero-stats">
          <div class="stat-card">
            <div class="stat-value">27</div>
            <div class="stat-label">Chapitres (15 Excel + 12 VBA) &amp; Projet NovaTech</div>
          </div>
          <div class="stat-card">
            <div class="stat-value">${ALL_OPERATIONS.length}</div>
            <div class="stat-label">Opérations sur 12 catégories + Scénarios métier</div>
          </div>
          <div class="stat-card">
            <div class="stat-value">9</div>
            <div class="stat-label">Formats d'exercices interactifs &amp; suivi adaptatif</div>
          </div>
          <div class="stat-card">
            <div class="stat-value">0</div>
            <div class="stat-label">Exécution automatique de macros (contrôle statique strict)</div>
          </div>
        </div>
      </section>

      <section class="stack-md">
        <div class="page-header">
          <div>
            <h2>Deux grands piliers complémentaires</h2>
            <p class="page-subtitle">Avancez à votre rythme, sans être obligé de savoir coder pour utiliser Excel ou automatiser vos tâches.</p>
          </div>
        </div>
        <div class="cards-grid-3">
          <article class="panel stack-sm">
            <span class="badge badge-success">Pilier 1 · Académie Excel</span>
            <h3>Parcours A (Excel) &amp; Parcours B (VBA)</h3>
            <p class="field-hint">15 chapitres pour maîtriser formules, TCD, tableaux structurés, Power Query et KPI, puis 12 chapitres VBA + 12 notions expertes autour de NovaTech.</p>
            <button type="button" class="btn-secondary btn-sm" data-nav="tutorial">Découvrir l'Académie</button>
          </article>
          <article class="panel stack-sm">
            <span class="badge badge-info">Pilier 2 · Assistant Intelligent</span>
            <h3>Du besoin métier au code adapté</h3>
            <p class="field-hint">Décrivez votre besoin : l'assistant le reformule, pose les bonnes questions et génère du VBA (Desktop) ou des Office Scripts (Excel Web).</p>
            <button type="button" class="btn-secondary btn-sm" data-nav="free">Ouvrir l'Assistant</button>
          </article>
          <article class="panel stack-sm">
            <span class="badge badge-warning">Audit &amp; Scénarios métier</span>
            <h3>Corriger l'existant et réutiliser</h3>
            <p class="field-hint">Collez une macro existante pour détecter bugs, .Select et risques avec diff avant/après, ou piochez dans 7 domaines métier.</p>
            <button type="button" class="btn-secondary btn-sm" data-nav="auditor">Auditer une macro</button>
          </article>
        </div>
      </section>

      <section class="stack-md">
        <div class="page-header">
          <div>
            <h2>Aperçu des catégories d'automatisation</h2>
            <p class="page-subtitle">Parcourez l'ensemble du catalogue pour tous les métiers.</p>
          </div>
          <button type="button" class="btn-primary btn-sm" data-nav="catalog">Voir les 12 catégories (${ALL_OPERATIONS.length} opérations)</button>
        </div>
        <div class="cards-grid-3">${catPreview}</div>
      </section>
    </div>
  `;

  container.querySelectorAll('[data-nav]').forEach((btn) => {
    btn.addEventListener('click', () => navigate(btn.dataset.nav));
  });

  container.querySelectorAll('[data-open-cat]').forEach((btn) => {
    btn.addEventListener('click', () => {
      setState({ catalogFilterCategory: btn.dataset.openCat }, 'catalogFilterCategory');
      navigate('catalog');
    });
  });
}
