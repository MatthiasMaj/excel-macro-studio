/**
 * Suivi pédagogique adaptatif persistant dans localStorage.
 * Enregistre la progression sur les deux parcours (Excel & VBA),
 * les résultats aux exercices/quiz, la maîtrise par compétence
 * et génère des recommandations de révision ciblées sans jamais bloquer l'apprenant.
 */

const ACADEMY_PROGRESS_KEY = 'ems-academy-progress-v3';
const LEGACY_PROGRESS_KEY = 'ems-curriculum-progress-v2';

// Fallback en mémoire lorsque localStorage est indisponible (tests Node, navigation privée stricte)
let memoryProgress = null;

function defaultProgress() {
  return {
    viewed: {},           // { [chapterId]: isoDate }
    completed: {},        // { [chapterId]: true }
    exercises: {},        // { [chapterId]: { [exKey]: { passed: bool, attempts: n, fails: n } } }
    quizzes: {},          // { [chapterId]: { bestScore: n, total: n, attempts: n } }
    competencies: {},     // { [competencyId]: { successes: n, fails: n } }
    lastVisited: null,    // { trackId: 'excel' | 'vba', chapterId: string, updatedAt: string }
    updatedAt: null
  };
}

function readLegacyCompleted() {
  try {
    if (typeof localStorage === 'undefined') return [];
    const raw = localStorage.getItem(LEGACY_PROGRESS_KEY);
    const parsed = raw ? JSON.parse(raw) : [];
    return Array.isArray(parsed) ? parsed : [];
  } catch {
    return [];
  }
}

export function loadAcademyProgress() {
  const base = defaultProgress();
  try {
    if (typeof localStorage !== 'undefined') {
      const raw = localStorage.getItem(ACADEMY_PROGRESS_KEY);
      if (raw) {
        const parsed = JSON.parse(raw);
        memoryProgress = { ...base, ...parsed };
        return memoryProgress;
      }
      // Migration douce de l'ancien historique v2
      const legacyIds = readLegacyCompleted();
      if (legacyIds.length > 0) {
        for (const id of legacyIds) {
          base.completed[id] = true;
          base.viewed[id] = new Date().toISOString();
        }
      }
    }
  } catch {
    return memoryProgress || base;
  }
  return memoryProgress ? { ...base, ...memoryProgress } : base;
}

export function saveAcademyProgress(progress) {
  const next = { ...progress, updatedAt: new Date().toISOString() };
  memoryProgress = next;
  try {
    if (typeof localStorage !== 'undefined') {
      localStorage.setItem(ACADEMY_PROGRESS_KEY, JSON.stringify(next));
      // Synchronise les chapitres validés avec l'ancienne clé pour rétrocompatibilité
      const completedIds = Object.keys(next.completed || {}).filter((k) => next.completed[k]);
      localStorage.setItem(LEGACY_PROGRESS_KEY, JSON.stringify(completedIds));
    }
  } catch {
    // Stockage bloqué ou quota atteint
  }
  return next;
}

export function markChapterViewed(trackId, chapterId) {
  const p = loadAcademyProgress();
  if (!p.viewed[chapterId]) {
    p.viewed[chapterId] = new Date().toISOString();
  }
  p.lastVisited = { trackId, chapterId, updatedAt: new Date().toISOString() };
  return saveAcademyProgress(p);
}

export function markChapterCompleted(trackId, chapterId, isCompleted = true) {
  const p = loadAcademyProgress();
  if (isCompleted) {
    p.completed[chapterId] = true;
    if (!p.viewed[chapterId]) p.viewed[chapterId] = new Date().toISOString();
  } else {
    delete p.completed[chapterId];
  }
  p.lastVisited = { trackId, chapterId, updatedAt: new Date().toISOString() };
  return saveAcademyProgress(p);
}

export function recordExerciseAttempt(chapterId, exerciseKey, passed, competencyId = null) {
  const p = loadAcademyProgress();
  if (!p.exercises[chapterId]) p.exercises[chapterId] = {};
  const prev = p.exercises[chapterId][exerciseKey] || { passed: false, attempts: 0, fails: 0 };
  p.exercises[chapterId][exerciseKey] = {
    passed: prev.passed || Boolean(passed),
    attempts: prev.attempts + 1,
    fails: prev.fails + (passed ? 0 : 1)
  };

  if (competencyId) {
    const comp = p.competencies[competencyId] || { successes: 0, fails: 0 };
    if (passed) comp.successes += 1;
    else comp.fails += 1;
    p.competencies[competencyId] = comp;
  }

  return saveAcademyProgress(p);
}

export function recordQuizScore(chapterId, score, total, competencyId = null) {
  const p = loadAcademyProgress();
  const prev = p.quizzes[chapterId] || { bestScore: 0, total, attempts: 0 };
  p.quizzes[chapterId] = {
    bestScore: Math.max(prev.bestScore || 0, score),
    total,
    lastScore: score,
    attempts: prev.attempts + 1
  };
  if (competencyId) {
    const comp = p.competencies[competencyId] || { successes: 0, fails: 0 };
    if (total > 0 && score / total >= 0.6) comp.successes += 1;
    else comp.fails += 1;
    p.competencies[competencyId] = comp;
  }
  return saveAcademyProgress(p);
}

export function resetAcademyProgress() {
  const fresh = defaultProgress();
  return saveAcademyProgress(fresh);
}

export function getTrackStats(chapters = [], progress = loadAcademyProgress()) {
  const total = chapters.length;
  const completedCount = chapters.filter((c) => Boolean(progress.completed?.[c.id])).length;
  const viewedCount = chapters.filter((c) => Boolean(progress.viewed?.[c.id])).length;
  let exercisesPassed = 0;
  let exercisesTotal = 0;

  for (const c of chapters) {
    const map = progress.exercises?.[c.id] || {};
    for (const k of ['guided', 'autonomous', 'remediation']) {
      if (k !== 'remediation') exercisesTotal += 1;
      if (map[k]?.passed) exercisesPassed += 1;
    }
  }

  const percent = total > 0 ? Math.round((completedCount / total) * 100) : 0;
  return {
    total,
    completedCount,
    viewedCount,
    exercisesPassed,
    exercisesTotal,
    percent
  };
}

export function getLevelValidationStatus(levels = [], chapters = [], progress = loadAcademyProgress()) {
  return levels.map((lvl) => {
    const lvlChapters = chapters.filter((c) => c.levelId === lvl.id);
    const done = lvlChapters.filter((c) => Boolean(progress.completed?.[c.id])).length;
    const total = lvlChapters.length;
    const validated = total > 0 && done === total;
    return {
      ...lvl,
      done,
      total,
      validated,
      percent: total > 0 ? Math.round((done / total) * 100) : 0
    };
  });
}

export function getCompetencyMasterySummary(progress = loadAcademyProgress()) {
  return Object.entries(progress.competencies || {}).map(([name, stat]) => {
    const attempts = (stat.successes || 0) + (stat.fails || 0);
    const ratio = attempts > 0 ? Math.round((stat.successes / attempts) * 100) : 0;
    const level =
      ratio >= 80 && stat.successes >= 2
        ? 'Maîtrisée'
        : ratio >= 50
          ? 'En consolidation'
          : 'À retravailler';
    return {
      competency: name,
      successes: stat.successes || 0,
      fails: stat.fails || 0,
      ratio,
      level
    };
  });
}

/**
 * Détecte si l'apprenant rencontre des difficultés sur un chapitre précis
 * afin de débloquer une remédiation ciblée ou une recommandation de révision.
 */
export function needsChapterRemediation(chapterId, progress = loadAcademyProgress()) {
  const exMap = progress.exercises?.[chapterId] || {};
  const failCount = Object.values(exMap).reduce((sum, e) => sum + (e.fails || 0), 0);
  const quiz = progress.quizzes?.[chapterId];
  const lowQuiz = quiz && quiz.total > 0 && quiz.lastScore !== undefined && quiz.lastScore / quiz.total < 0.6;
  return failCount >= 2 || Boolean(lowQuiz);
}

export function getReviewRecommendations(allChapters = [], progress = loadAcademyProgress()) {
  const recs = [];
  for (const ch of allChapters) {
    if (needsChapterRemediation(ch.id, progress)) {
      const quiz = progress.quizzes?.[ch.id];
      const exMap = progress.exercises?.[ch.id] || {};
      const failCount = Object.values(exMap).reduce((sum, e) => sum + (e.fails || 0), 0);
      recs.push({
        chapterId: ch.id,
        trackId: ch.trackId || 'vba',
        title: ch.title,
        levelName: ch.levelName,
        reason:
          quiz && quiz.total > 0 && quiz.lastScore / quiz.total < 0.6
            ? `Score de validation à renforcer (${quiz.lastScore}/${quiz.total})`
            : `${failCount} erreur(s) relevée(s) sur les exercices pratiques — exercice de consolidation disponible.`
      });
    }
  }
  return recs;
}
