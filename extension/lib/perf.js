// Wild Themes: smoothness guard maths, shared by the content script and the tests.
// gaps: milliseconds between consecutive animation frames while the scene runs. Returns 1 when the
// frames are late or uneven (this device can't keep up: the scene should move less), else 0.
self.WWPerf = {
  judge(gaps, ms) {
    if (!gaps || gaps.length < ms / 1000 * 10) return 1;                 // under ~10 frames a second
    const s = gaps.slice().sort((a, b) => a - b), median = s[s.length >> 1];
    const slow = gaps.filter((g) => g > 40).length / gaps.length;         // share of frames that took 40 ms or more
    return median > 28 || slow > .12 ? 1 : 0;
  },
};
