// Wild Themes — small UI pieces shared by the website's file studio, the extension's file page
// and the Office add-in, so every tool shows a file's theme and its settings the same way.
// Needs wtcore.js. No styling here: each tool styles the classes (.badge, .switch, .field, .look h2).
(() => {
  const WT = self.WTCore;
  const LABEL = { none: "No theme", intact: "Intact", modified: "Changed", partial: "Incomplete", legacy: "Older format", damaged: "Damaged", unverified: "Unchecked" };
  const KIND = { docx: "Word document", pptx: "PowerPoint deck", xlsx: "Excel workbook" };
  const GROUPS = [["Look", ["backdrop", "panel", "styles", "tables", "lists", "direct", "tabs", "lock"]], ["Effects", ["animated", "transition", "glow"]]];

  // Settings form for one file type. values: current settings; onEdit(key, value).
  function settingsForm(form, kind, values, onEdit) {
    form.textContent = "";
    for (const [title, keys] of GROUPS) {
      const ks = keys.filter((k) => WT.FEATURES[kind].includes(k));
      if (!ks.length) continue;
      const h = document.createElement("h2"); h.textContent = title; form.append(h);
      for (const k of ks) {
        const spec = WT.SETTINGS[k];
        if (spec.type === "bool") {
          const l = document.createElement("label"), i = document.createElement("input"), sp = document.createElement("span"), b = document.createElement("b"), sm = document.createElement("small");
          l.className = "switch"; i.type = "checkbox"; i.checked = !!values[k]; i.onchange = () => onEdit(k, i.checked);
          b.textContent = spec.label; sm.textContent = spec.tip; sp.append(b, sm); l.append(i, sp); form.append(l);
        } else {
          const d = document.createElement("div"), lab = document.createElement("label"), sm = document.createElement("small");
          d.className = "field"; lab.textContent = spec.label; sm.textContent = spec.tip;
          let input;
          if (spec.type === "range") {
            input = Object.assign(document.createElement("input"), { type: "range", min: spec.min, max: spec.max, step: 5, value: values[k] });
            input.oninput = () => onEdit(k, Number(input.value));
          } else {
            input = document.createElement("select");
            for (const o of spec.options) input.append(new Option(o[0].toUpperCase() + o.slice(1), o));
            input.value = values[k]; input.onchange = () => onEdit(k, input.value);
          }
          input.id = "wt-set-" + k; lab.htmlFor = input.id;
          d.append(lab, input, sm); form.append(d);
        }
      }
    }
  }
  // Human summary of WTCore.inspect() output.
  function describe(info) {
    if (!info) return { title: "Couldn't read this file", sub: "", badge: "", cls: "" };
    const st = info.status, t = info.theme;
    const when = info.applied ? new Date(info.applied).toLocaleDateString() : "";
    return {
      title: st === "none" ? `No theme in this ${KIND[info.kind] || "file"}` : (t.name || t.id),
      sub: st === "none" ? "Pick a theme to save one into it." : [KIND[info.kind], when && `saved ${when}`, info.tool && `with ${info.tool}`].filter(Boolean).join(" · "),
      badge: st === "none" ? "" : LABEL[st] || st, cls: st,
      problems: (info.problems || []).map((p) => p.text),
      canRepair: ["modified", "partial", "legacy", "damaged"].includes(st), canRemove: st !== "none",
    };
  }
  const swatch = (t) => `linear-gradient(135deg,${(t && t.accent) || "#444"},${(t && (t.accent2 || t.accent)) || "#444"})`;
  self.WTUI = { settingsForm, describe, swatch, LABEL, KIND };
})();
