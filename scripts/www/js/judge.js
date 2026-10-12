/**
 * judge.js（浏览器版） — 判分核心（纯函数，可独立 Node 测试）
 * 判分规则与 parser-core 的 answerRule 对应：
 *   exact     单选：用户选项 key 与答案相等
 *   set       多选：选项集合相等（顺序无关）
 *   yn        判断：'Y'/'N' 相等
 *   multiword 填空：逐空 trim 后相等（大小写不敏感）
 *   self      简答：用户自评（前端传 'known' | 'vague' | 'unknown'）
 */
function normText(v) {
  if (v === undefined || v === null) return '';
  return String(v).trim();
}

function judge(question, userAnswer) {
  const rule = question.answerRule;
  const ans = question.answer;

  switch (rule) {
    case 'exact':
      return { correct: normText(userAnswer).toUpperCase() === normText(ans).toUpperCase() };

    case 'set': {
      const a = (Array.isArray(ans) ? ans : []).slice().map((x) => normText(x).toUpperCase()).sort().join('');
      const b = (Array.isArray(userAnswer) ? userAnswer : []).slice().map((x) => normText(x).toUpperCase()).sort().join('');
      return { correct: a.length > 0 && a === b };
    }

    case 'yn':
      return { correct: normText(userAnswer).toUpperCase() === normText(ans).toUpperCase() };

    case 'multiword': {
      const expect = Array.isArray(ans) ? ans : [ans];
      const given = Array.isArray(userAnswer) ? userAnswer : [userAnswer];
      if (expect.length !== given.length) return { correct: false };
      const ok = expect.every((x, i) =>
        normText(x).replace(/\s+/g, '').toLowerCase() === normText(given[i]).replace(/\s+/g, '').toLowerCase()
      );
      return { correct: ok };
    }

    case 'self':
      // 自评：known 计对，vague/unknown 计错
      return { correct: userAnswer === 'known', selfRated: true };

    default:
      return { correct: false, error: 'unknown rule: ' + rule };
  }
}

/** 间隔重复：答错归零，之后依次 1 天 → 3 天 → 7 天 → 15 天（毫秒） */
const REVIEW_INTERVALS = [
  1 * 24 * 3600 * 1000,   // stage 0 → 1 天后
  3 * 24 * 3600 * 1000,   // stage 1 → 3 天后
  7 * 24 * 3600 * 1000,   // stage 2 → 7 天后
  15 * 24 * 3600 * 1000   // stage 3 → 15 天后
];
const MASTERED_STAGE = 2; // 从 0 起连续答对 3 次（stage 到 2 之上）视为掌握

function nextStage(curStage, correct) {
  if (!correct) return 0;
  return Math.min(curStage + 1, REVIEW_INTERVALS.length);
}

window.Judge = {
  judge: judge,
  nextStage: nextStage,
  REVIEW_INTERVALS: REVIEW_INTERVALS,
  MASTERED_STAGE: MASTERED_STAGE
};
