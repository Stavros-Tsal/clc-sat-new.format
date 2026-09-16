// ═══════════════════════════════════════════════════════════════════
// CLC SAT — Shared Progress Tracking
// Client-side only: everything is stored in this browser's localStorage.
// Nothing is sent anywhere. The student can export a report (PDF via
// print) and send it to the teacher themselves.
// ═══════════════════════════════════════════════════════════════════
(function () {
  const KEY = 'clc_progress_v1';

  function defaultProgress() {
    return {
      drills: { "0": {}, "1": {}, "2": {}, "3": {} },
      drillStats: {
        "0": { correct: 0, total: 0 },
        "1": { correct: 0, total: 0 },
        "2": { correct: 0, total: 0 },
        "3": { correct: 0, total: 0 }
      },
      mathTests: {},
      mathTestStats: { correct: 0, total: 0 },
      fullTests: {},
      fullTestStats: { correct: 0, total: 0 },
      specificPractice: {}
    };
  }

  function load() {
    let data;
    try {
      const raw = localStorage.getItem(KEY);
      data = raw ? JSON.parse(raw) : defaultProgress();
    } catch (e) {
      data = defaultProgress();
    }
    const def = defaultProgress();
    Object.keys(def).forEach(function (k) {
      if (!data[k]) data[k] = def[k];
    });
    ["0", "1", "2", "3"].forEach(function (t) {
      if (!data.drills[t]) data.drills[t] = {};
      if (!data.drillStats[t]) data.drillStats[t] = { correct: 0, total: 0 };
    });
    return data;
  }

  function save(data) {
    try { localStorage.setItem(KEY, JSON.stringify(data)); } catch (e) { /* storage full/unavailable — ignore */ }
  }

  // Record one answered drill question. Attempts only accumulate until the
  // question is first answered correctly ("attempts to first success").
  window.clcRecordDrillAnswer = function (tabIdx, qIdx, isCorrect) {
    const p = load();
    const tabKey = String(tabIdx);
    if (!p.drills[tabKey]) p.drills[tabKey] = {};
    const qKey = String(qIdx);
    let rec = p.drills[tabKey][qKey];
    if (!rec) { rec = { attempts: 0, solved: false }; p.drills[tabKey][qKey] = rec; }
    if (!rec.solved) {
      rec.attempts += 1;
      if (isCorrect) rec.solved = true;
    }
    if (!p.drillStats[tabKey]) p.drillStats[tabKey] = { correct: 0, total: 0 };
    p.drillStats[tabKey].total += 1;
    if (isCorrect) p.drillStats[tabKey].correct += 1;
    save(p);
  };

  // Record one completed Math Practice Test attempt (Test 1–15, scaled /800).
  window.clcRecordMathTestAttempt = function (testId, scaledScore, correctCount, totalCount) {
    const p = load();
    const key = String(testId);
    if (!p.mathTests[key]) p.mathTests[key] = [];
    p.mathTests[key].push({ score: scaledScore, date: new Date().toISOString() });
    p.mathTestStats.correct += correctCount;
    p.mathTestStats.total += totalCount;
    save(p);
  };

  // Record one completed Full Practice Test attempt (Test 1–3, R&W+Math).
  // scaledScore may be null if the test only reports raw correct/total.
  window.clcRecordFullTestAttempt = function (testId, scaledScore, correctCount, totalCount) {
    const p = load();
    const key = String(testId);
    if (!p.fullTests[key]) p.fullTests[key] = [];
    p.fullTests[key].push({
      scaledScore: (scaledScore === null || scaledScore === undefined) ? null : scaledScore,
      correctCount: correctCount,
      totalCount: totalCount,
      date: new Date().toISOString()
    });
    p.fullTestStats.correct += correctCount;
    p.fullTestStats.total += totalCount;
    save(p);
  };

  // Record one completed Specific Practice attempt (Advanced Math / Algebra / PSDA).
  window.clcRecordSpecificAttempt = function (bankName, correctCount, totalCount) {
    const p = load();
    if (!p.specificPractice[bankName]) p.specificPractice[bankName] = [];
    p.specificPractice[bankName].push({
      correctCount: correctCount,
      totalCount: totalCount,
      date: new Date().toISOString()
    });
    save(p);
  };

  window.clcLoadProgress = load;
})();
