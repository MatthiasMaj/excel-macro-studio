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
          <span class="hero-badge">100 % en français · Hors-ligne pour le catalogue · VBA propre</span>
          <h1 class="page-title">Créez des macros Excel VBA fiables, sans savoir programmer</h1>
          <p class="page-subtitle">
            Assistant universel d'automatisation Excel pour particuliers, assistants, comptables, RH, gestionnaires, ingénieurs, responsables de production et analystes de données.
          </p>
          <div class="hero-actions">
            <button type="button" class="btn-primary" data-nav="guided">Démarrer le mode guidé</button>
            <button type="button" class="btn-secondary" data-nav="workflow">Créer un workflow multi-opérations</button>
            <button type="button" class="btn-secondary" data-nav="free">Décrire en langage naturel</button>
            <button type="button" class="btn-secondary" data-nav="tutorial">Suivre le cursus Excel &amp; VBA (12 chapitres)</button>
            <button type="button" class="btn-ghost" data-nav="analyzer">Analyser un fichier Excel</button>
          </div>
        </div>
        <div class="hero-stats">
          <div class="stat-card">
            <div class="stat-value">12</div>
            <div class="stat-label">Catégories métiers complètes (A à L)</div>
          </div>
          <div class="stat-card">
            <div class="stat-value">${ALL_OPERATIONS.length}</div>
            <div class="stat-label">Opérations prêtes à paramétrer</div>
          </div>
          <div class="stat-card">
            <div class="stat-value">3</div>
            <div class="stat-label">Modes : Guidé, Multi-étapes, Libre</div>
          </div>
          <div class="stat-card">
            <div class="stat-value">0</div>
            <div class="stat-label">Exécution automatique de macros (sécurité stricte)</div>
          </div>
        </div>
      </section>

      <section class="stack-md">
        <div class="page-header">
          <div>
            <h2>Comment fonctionne Excel Macro Studio ?</h2>
            <p class="page-subtitle">Trois méthodes complémentaires selon votre besoin et votre niveau.</p>
          </div>
        </div>
        <div class="cards-grid-3">
          <article class="panel stack-sm">
            <span class="badge badge-success">Mode 1 · Débutants &amp; Rapide</span>
            <h3>Génération guidée par formulaire</h3>
            <p class="field-hint">Choisissez une opération, remplissez des champs clairs (alimentés par votre fichier Excel si importé) et récupérez le VBA commenté.</p>
            <button type="button" class="btn-secondary btn-sm" data-nav="guided">Ouvrir le mode guidé</button>
          </article>
          <article class="panel stack-sm">
            <span class="badge badge-info">Mode 2 · Workflows complets</span>
            <h3>Générateur multi-opérations</h3>
            <p class="field-hint">Enchaînez plusieurs actions (ex. Importer CSV → Supprimer vides → Doublons → Tri → Graphique → PDF) dans un module unique.</p>
            <button type="button" class="btn-secondary btn-sm" data-nav="workflow">Construire un workflow</button>
          </article>
          <article class="panel stack-sm">
            <span class="badge badge-warning">Mode 3 · Langage naturel</span>
            <h3>Description libre</h3>
            <p class="field-hint">Décrivez votre besoin en français. Le moteur identifie les opérations, demande les précisions manquantes et génère le code.</p>
            <button type="button" class="btn-secondary btn-sm" data-nav="free">Essayer le mode libre</button>
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

