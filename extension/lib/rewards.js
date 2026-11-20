// Wild Themes — words typed, milestones and the secret theme. Shared by the page script (counts what you type), the
// background (keeps the totals and rolls the dice) and the popup (shows them).
// Every word you type while a theme is showing is one ticket: it has a 1 in ODDS chance to unlock the secret theme.
// Only real key presses count, one finished word at a time, so pasted, dropped or dictated text never does.
self.WWRewards = {
  ODDS: 37500,
  SECRET: "legendary",
  // passed by total words typed (all themes together); each one plays a short cutscene
  MILESTONES: [
    { at: 1000, id: "w1k", name: "Warmed up", emoji: "🔥" },
    { at: 10000, id: "w10k", name: "Wordsmith", emoji: "✒️" },
    { at: 50000, id: "w50k", name: "Novelist", emoji: "📚" },
    { at: 67000, id: "w67k", name: "Six seven", emoji: "🤲", cutscene: "67" },   // its own cutscene: "67" all over the screen
    { at: 100000, id: "w100k", name: "Keyboard legend", emoji: "🏆" },
  ],
  // What a key press does to the word being typed: "char" (a letter or digit: inside a word), "end" (a space, Enter, Tab or
  // punctuation: the word, if there is one, is finished), or "" (ignored: shortcuts such as Ctrl+V, held-down repeats,
  // keys that type nothing, text being composed by an input method, and events not made by a person).
  // An apostrophe or a hyphen inside a word keeps it one word ("don't", "well-known").
  keyKind(e) {
    if (!e || e.isTrusted === false || e.repeat || e.isComposing || e.metaKey) return "";
    if ((e.ctrlKey || e.altKey) && !(e.getModifierState && e.getModifierState("AltGraph"))) return "";   // AltGr types letters on many keyboards
    const k = e.key || "";
    if (k == "Enter" || k == "Tab") return "end";
    if ([...k].length != 1) return "";
    if (/[\p{L}\p{N}]/u.test(k)) return "char";
    if (k == "'" || k == "’" || k == "-") return "";
    return /[\s\p{P}\p{S}]/u.test(k) ? "end" : "";
  },
  total: (words) => Object.values(words || {}).reduce((a, b) => a + (+b || 0), 0),
  passed(total) { return this.MILESTONES.filter((m) => total >= m.at); },
  next(total) { return this.MILESTONES.find((m) => total < m.at) || null; },
};
