// Wild Themes — reward cutscenes, played over the page when you find the secret theme or pass a word milestone.
// Plain elements moved by the Web Animations API: no script runs per frame. They never take the keyboard, so typing carries
// on into the document. Esc or a click skips; they end by themselves. With reduced motion (or Animation: Still) the last
// frame simply fades in and out.
self.WWCut = (() => {
  const D = document;
  let cur = null;
  const el = (tag, cls, text) => { const n = D.createElement(tag); if (cls) n.className = cls; if (text != null) n.textContent = text; return n; };
  const fmt = (n) => Number(n || 0).toLocaleString();

  function stop(fast) {
    if (!cur) return;
    const c = cur; cur = null;
    clearTimeout(c.timer);
    removeEventListener("keydown", c.key, true);
    if (fast) return c.root.remove();
    c.root.style.pointerEvents = "none";
    c.root.animate([{ opacity: 1 }, { opacity: 0 }], { duration: 450, easing: "ease-in", fill: "forwards" }).onfinish = () => c.root.remove();
  }
  // jump to the end of everything that has an end (the rays keep turning)
  const skip = () => { if (!cur) return false; let moved = false; for (const a of cur.anims) if (a.playState != "finished" && isFinite(a.effect.getComputedTiming().endTime)) { a.finish(); moved = true; } return moved; };
  // Esc: first skips to the end, then closes
  const esc = () => { if (cur && !skip()) stop(); };

  function start(kind, accent, accent2, life) {
    stop(true);
    const root = el("div"); root.id = "ww-cut"; root.dataset.kind = kind;
    root.style.setProperty("--g", accent || "#ffd23f"); root.style.setProperty("--h", accent2 || accent || "#b388ff");
    const key = (e) => { if (e.key == "Escape") esc(); };
    addEventListener("keydown", key, true);
    cur = { root, anims: [], key, timer: setTimeout(() => stop(), life) };
    return cur;
  }
  const play = (c, node, frames, delay, duration, easing = "ease-out", more = {}) => {
    const a = node.animate(frames, { delay, duration, easing, fill: "both", ...more });
    c.anims.push(a); return a;
  };
  // a ring of sparks and confetti thrown out from the middle of the box r (measured before anything moved)
  function burst(c, r, n, delay, reach) {
    const b = el("div", "wwc-burst");
    for (let i = 0; i < n; i++) b.append(el("i"));
    c.root.append(b);
    b.style.left = r.left + r.width / 2 + "px"; b.style.top = r.top + r.height / 2 + "px";
    [...b.children].forEach((p, i) => {
      const a = (i / n) * 6.283 + Math.random() * .4, d = reach * (.55 + Math.random() * .6), spin = (Math.random() * 720 - 360) | 0;
      play(c, p, [{ opacity: 1, transform: "translate(0,0) scale(.4)" }, { opacity: 1, offset: .7 },
        { opacity: 0, transform: `translate(${Math.cos(a) * d}px,${Math.sin(a) * d + reach * .25}px) rotate(${spin}deg) scale(1)` }],
      delay + Math.random() * 120, 1100 + Math.random() * 700, "cubic-bezier(.15,.7,.3,1)");
    });
  }

  // The big one: a locked card shakes, breaks open in a flash and turns over to show the secret theme.
  function secret({ name, emoji, accent, accent2, words, odds, calm, use }) {
    const c = start("secret", accent, accent2, 16000), R = c.root;
    R.setAttribute("role", "dialog"); R.setAttribute("aria-label", `Secret theme unlocked: ${name}`);
    const shade = el("i", "wwc-shade"), top = el("i", "wwc-bar wwc-top"), bot = el("i", "wwc-bar wwc-bot"), flash = el("i", "wwc-flash");
    const stage = el("div", "wwc-stage"), rays = el("i", "wwc-rays"), pop = el("div", "wwc-pop"), shake = el("div", "wwc-shake"), card = el("div", "wwc-card");
    const back = el("div", "wwc-face wwc-back"), front = el("div", "wwc-face wwc-front");
    back.append(el("b", "wwc-q", "?"), el("small", "", "LOCKED"));
    front.append(el("b", "wwc-emoji", emoji || "👑"), el("small", "", name));
    card.append(back, front); shake.append(card); pop.append(shake);
    const kicker = el("p", "wwc-kicker", "Secret theme unlocked"), title = el("h2", "wwc-title", name);
    const sub = el("p", "wwc-sub", `Each typed word had a 1 in ${fmt(odds)} chance. You found it after ${fmt(words)} words.`);
    const btns = el("div", "wwc-btns"), go = el("button", "wwc-go", "Use it now"), later = el("button", "wwc-later", "Later");
    go.type = later.type = "button"; btns.append(go, later);
    stage.append(pop, kicker, title, sub, btns);
    R.append(shade, top, bot, rays, stage, flash);
    D.documentElement.append(R);
    go.onclick = (e) => { e.stopPropagation(); use && use(); stop(); };
    later.onclick = (e) => { e.stopPropagation(); stop(); };
    R.onclick = () => esc();
    const cr = pop.getBoundingClientRect();
    rays.style.left = cr.left + cr.width / 2 + "px"; rays.style.top = cr.top + cr.height / 2 + "px";

    play(c, R, [{ opacity: 0 }, { opacity: 1 }], 0, 500);
    play(c, top, [{ transform: "translateY(-100%)" }, { transform: "none" }], 0, 600, "cubic-bezier(.2,.8,.2,1)");
    play(c, bot, [{ transform: "translateY(100%)" }, { transform: "none" }], 0, 600, "cubic-bezier(.2,.8,.2,1)");
    play(c, pop, [{ opacity: 0, transform: "scale(.55) translateY(40px)" }, { opacity: 1, transform: "none" }], 350, 650, "cubic-bezier(.2,1.4,.4,1)");
    // the lock strains: shaking harder and harder, its glow pulsing faster
    const wob = [0, -2, 2, -3, 3, -4, 4, -6, 6, -8, 8, 0].map((d, i, a) => ({ transform: `rotate(${d}deg) scale(${1 + i / a.length * .1})` }));
    play(c, shake, wob, 1000, 1300, "linear");
    play(c, back, [{ boxShadow: "0 0 0 2px #fff2,0 0 0 var(--g)" }, { boxShadow: "0 0 0 2px var(--g),0 0 70px var(--g)" }], 1000, 1300, "ease-in");
    play(c, rays, [{ opacity: 0, transform: "rotate(0deg) scale(.6)" }, { opacity: .42, transform: "rotate(40deg) scale(1)" }], 1200, 1300, "ease-out");
    c.anims.push(rays.animate([{ rotate: "0deg" }, { rotate: "360deg" }], { duration: 24000, iterations: Infinity }));
    // it breaks open: a flash, the card turns over, sparks fly
    play(c, flash, [{ opacity: 0 }, { opacity: .95, offset: .25 }, { opacity: 0 }], 2300, 900, "ease-out");
    play(c, card, [{ transform: "rotateY(0deg)" }, { transform: "rotateY(180deg)" }], 2350, 800, "cubic-bezier(.3,1.3,.5,1)");
    play(c, pop, [{ transform: "scale(1)" }, { transform: "scale(1.12)" }, { transform: "scale(1)" }], 2350, 700, "ease-out", { composite: "add" });
    burst(c, cr, 30, 2400, Math.max(160, Math.min(innerWidth, innerHeight) * .42));
    // the title card
    play(c, kicker, [{ opacity: 0, letterSpacing: ".9em" }, { opacity: 1, letterSpacing: ".28em" }], 2700, 900);
    play(c, title, [{ opacity: 0, transform: "scale(1.9)" }, { opacity: 1, transform: "scale(1)" }], 2950, 650, "cubic-bezier(.2,1.2,.4,1)");
    play(c, sub, [{ opacity: 0, transform: "translateY(12px)" }, { opacity: 1, transform: "none" }], 3500, 600);
    play(c, btns, [{ opacity: 0, transform: "translateY(12px)" }, { opacity: 1, transform: "none" }], 3900, 600);
    if (calm) skip();
    return R;
  }

  // A milestone: a medal drops in, spins and shows how far you've come, then floats away. Doesn't block anything.
  function badge({ name, emoji, at, total, accent, accent2, calm }) {
    const c = start("badge", accent, accent2, 4600), R = c.root;
    R.setAttribute("role", "status");
    const stage = el("div", "wwc-stage"), medal = el("div", "wwc-medal"), ribbon = el("div", "wwc-ribbon");
    medal.append(el("b", "", emoji));
    ribbon.append(el("b", "", name), el("small", "", `${fmt(at)} words typed${total > at ? ` (${fmt(total)} so far)` : ""}`));
    stage.append(medal, ribbon); R.append(stage);
    D.documentElement.append(R);
    const mr = medal.getBoundingClientRect();
    play(c, medal, [{ opacity: 0, transform: "translateY(-60vh) rotateY(0deg)" }, { opacity: 1, transform: "translateY(0) rotateY(540deg)", offset: .7 },
      { transform: "translateY(-14px) rotateY(660deg)", offset: .85 }, { opacity: 1, transform: "translateY(0) rotateY(720deg)" }], 0, 1300, "ease-out");
    play(c, ribbon, [{ opacity: 0, transform: "scaleX(.2)" }, { opacity: 1, transform: "none" }], 1000, 450, "cubic-bezier(.2,1.3,.4,1)");
    burst(c, mr, 18, 950, 120);
    if (calm) skip();                                                // still: the medal and ribbon are simply there, then fade
    play(c, stage, [{ opacity: 1, transform: "none" }, { opacity: 0, transform: calm ? "none" : "translateY(-40px)" }], 3800, 700, "ease-in");
    return R;
  }
  return { secret, badge, esc, stop: () => stop(true), playing: () => !!cur };
})();
