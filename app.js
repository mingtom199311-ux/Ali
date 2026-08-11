(() => {
  "use strict";

  const stage = document.getElementById("stage");
  const slides = Array.from(document.querySelectorAll(".slide"));
  const total = slides.length;
  let currentIndex = Math.max(0, Math.min(total - 1, Number(location.hash.slice(1)) - 1 || 0));
  let transitioning = false;
  let transitionTimer = null;
  let touchStartX = null;
  let toastTimer = null;
  let audioContext = null;
  let soundEnabled = true;

  const currentNumber = document.getElementById("currentNumber");
  const totalNumber = document.getElementById("totalNumber");
  const sectionPill = document.getElementById("sectionPill");
  const progressBar = document.getElementById("progressBar");
  const outlineDrawer = document.getElementById("outlineDrawer");
  const notesDrawer = document.getElementById("notesDrawer");
  const outlineList = document.getElementById("outlineList");
  const notesTitle = document.getElementById("notesTitle");
  const notesTime = document.getElementById("notesTime");
  const notesContent = document.getElementById("notesContent");
  const toast = document.getElementById("toast");
  const soundButton = document.getElementById("toggleSound");
  const coachPanel = document.getElementById("coachPanel");
  const coachPrompt = document.getElementById("coachPrompt");
  const coachStatus = document.getElementById("coachStatus");
  const coachSkip = document.getElementById("coachSkip");
  const coachNext = document.getElementById("coachNext");

  const taskConfigs = [
    { selector: ".start-button", prompt: "点击“进入小教室”，开始今天的六步训练", done: "准备好啦，跟阿狸老师出发！" },
    { selector: ".foundation-labels article", prompt: "依次点击四张小卡，点亮稳定作文的地基", done: "四块地基已经点亮：完整、准确、规范、清楚。" },
    { selector: ".track-card", prompt: "点击“小作文”和“大作文”，完成一次对比", done: "两条赛道都看过了：先判断任务，再决定结构。" },
    { selector: ".topic-pills article", prompt: "点击三种题型，看清三条不同的写作路线", done: "题型决定路线，审题时先完成这次分流。" },
    { selector: ".score-tab", prompt: "切换小作文与大作文，查看两套评分档位", done: "评分不是黑箱：先守住档位，再减少扣分。" },
    { selector: ".deduction-grid article", prompt: "逐张点击扣分卡，揭晓最常见的失分代价", done: "五类扣分都已揭晓，先堵住最容易丢的分。" },
    { selector: ".big-list li", prompt: "点击四项训练，收集阿狸老师的备考工具", done: "四件工具已收齐：词块、句型、段落和时态。" },
    { selector: ".habit-path article", prompt: "按箭头顺序点击，走完五个高分习惯", done: "习惯链已走通，考场上按同一顺序执行。", sequential: true },
    { selector: ".rule-tab", prompt: "点击五个标签，检查稿纸里的五条规则", done: "五条格式规则已复核，卷面规范会直接保护分数。" },
    { selector: ".six-chips span", prompt: "从左到右点击六步，把流程装进脑子里", done: "六步路线已走完：定类、圈点、分段、打稿、查错、誊清。", sequential: true },
    { selector: ".error-item", prompt: "点击四句错例，亲手把它们改正确", done: "四个易错点已修正，正确形式要整块记住。" },
    { selector: ".opt-grid article", prompt: "点击三张优化卡，让表达一步步变清楚", done: "三次优化完成：句子更短，逻辑更清楚，敬体更稳定。" },
    { selector: "#checkGrid button", prompt: "按顺序点击六组检查项，完成考场复核", done: "六组检查全部完成，现在可以安心誊清。", sequential: true },
    { selector: ".final-six b", prompt: "按顺序点亮六个动作，完成最后一次复述", done: "恭喜通关！这六步就是你的稳定作文路线。", sequential: true }
  ];

  const taskStates = taskConfigs.map((config, index) => ({
    config,
    targets: Array.from(slides[index].querySelectorAll(config.selector)),
    completed: new Set()
  }));

  function fitStage() {
    const scale = Math.min(window.innerWidth / 1600, window.innerHeight / 900);
    document.documentElement.style.setProperty("--stage-scale", scale.toFixed(4));
  }

  function playSound(kind = "click", force = false) {
    if (!soundEnabled && !force) return;
    const AudioCtor = window.AudioContext || window.webkitAudioContext;
    if (!AudioCtor) return;
    if (!audioContext) audioContext = new AudioCtor();
    if (audioContext.state === "suspended") audioContext.resume();
    const sequences = {
      page: [[540, 0, .055], [720, .055, .07]],
      back: [[620, 0, .05], [440, .05, .07]],
      reveal: [[780, 0, .045]],
      click: [[690, 0, .04]],
      success: [[660, 0, .05], [880, .052, .08]],
      off: [[520, 0, .05], [360, .05, .07]]
    };
    const notes = sequences[kind] || sequences.click;
    const now = audioContext.currentTime;
    notes.forEach(([frequency, delay, duration]) => {
      const oscillator = audioContext.createOscillator();
      const gain = audioContext.createGain();
      oscillator.type = kind === "page" || kind === "success" ? "sine" : "triangle";
      oscillator.frequency.setValueAtTime(frequency, now + delay);
      oscillator.frequency.exponentialRampToValueAtTime(frequency * .92, now + delay + duration);
      gain.gain.setValueAtTime(.0001, now + delay);
      gain.gain.exponentialRampToValueAtTime(.032, now + delay + .008);
      gain.gain.exponentialRampToValueAtTime(.0001, now + delay + duration);
      oscillator.connect(gain).connect(audioContext.destination);
      oscillator.start(now + delay);
      oscillator.stop(now + delay + duration + .01);
    });
  }

  function showToast(message) {
    clearTimeout(toastTimer);
    toast.textContent = message;
    toast.classList.add("is-visible");
    toastTimer = setTimeout(() => toast.classList.remove("is-visible"), 1700);
  }

  function setDrawer(drawer, open) {
    drawer.classList.toggle("is-open", open);
    drawer.setAttribute("aria-hidden", String(!open));
  }

  function closeDrawers() {
    setDrawer(outlineDrawer, false);
    setDrawer(notesDrawer, false);
  }

  function taskIsComplete(index) {
    const state = taskStates[index];
    return state.targets.length > 0 && state.completed.size >= state.targets.length;
  }

  function updateTargetHints(index) {
    const state = taskStates[index];
    state.targets.forEach(target => target.classList.remove("is-suggested"));
    if (taskIsComplete(index)) return;
    const nextTarget = state.targets.find((target, position) => !state.completed.has(position));
    if (nextTarget) nextTarget.classList.add("is-suggested");
  }

  function updateCoach() {
    const state = taskStates[currentIndex];
    const complete = taskIsComplete(currentIndex);
    const done = state.completed.size;
    const nextSlideButton = document.getElementById("nextSlide");
    coachPanel.classList.toggle("is-complete", complete);
    coachPrompt.textContent = complete ? state.config.done : state.config.prompt;
    coachStatus.textContent = complete ? `完成 ${done} / ${state.targets.length}` : `请点击 · ${done} / ${state.targets.length}`;
    coachSkip.hidden = complete;
    coachNext.disabled = !complete;
    coachNext.textContent = currentIndex === total - 1 && complete
      ? "重新体验 ↺"
      : complete
        ? `进入第 ${String(currentIndex + 2).padStart(2, "0")} 页 →`
        : "完成后继续";
    coachNext.setAttribute("aria-label", complete ? coachNext.textContent : "请先完成本页点击任务");
    nextSlideButton.classList.toggle("is-ready", complete);
    nextSlideButton.title = complete ? "本页已完成，进入下一页" : "请先完成本页点击任务";
    updateTargetHints(currentIndex);
  }

  function markTask(index, position) {
    const state = taskStates[index];
    if (state.completed.has(position)) {
      playSound("click");
      return;
    }
    state.completed.add(position);
    state.targets[position].classList.add("is-done");
    slides[index].classList.add("is-task-started");
    const complete = taskIsComplete(index);
    if (complete) slides[index].classList.add("is-task-complete");
    playSound(complete ? "success" : "reveal");
    if (index === currentIndex) {
      updateCoach();
      if (complete) {
        coachNext.classList.remove("is-inviting");
        void coachNext.offsetWidth;
        coachNext.classList.add("is-inviting");
      }
    }
  }

  function handleTaskActivation(index, position, event) {
    const state = taskStates[index];
    if (state.config.sequential && !state.completed.has(position) && position !== state.completed.size) {
      event.preventDefault();
      event.stopImmediatePropagation();
      coachPanel.classList.remove("needs-attention");
      void coachPanel.offsetWidth;
      coachPanel.classList.add("needs-attention");
      showToast(`先点击第 ${state.completed.size + 1} 步，跟着发光提示走`);
      playSound("back");
      return;
    }
    markTask(index, position);
  }

  function bindGuidedTasks() {
    taskStates.forEach((state, slideIndex) => {
      state.targets.forEach((target, position) => {
        target.classList.add("task-target");
        target.dataset.taskPosition = String(position);
        if (target.tagName !== "BUTTON") {
          target.setAttribute("role", "button");
          target.tabIndex = 0;
        }
        target.addEventListener("click", event => handleTaskActivation(slideIndex, position, event));
        target.addEventListener("keydown", event => {
          if (event.key === "Enter" || event.key === " ") {
            event.preventDefault();
            target.click();
          }
        });
      });
    });

    coachSkip.addEventListener("click", () => {
      const state = taskStates[currentIndex];
      state.targets.forEach((target, position) => {
        state.completed.add(position);
        target.classList.add("is-done");
      });
      slides[currentIndex].classList.add("is-task-started", "is-task-complete");
      updateCoach();
      playSound("success");
      showToast("本页已标记完成，可以继续");
    });

    coachNext.addEventListener("click", event => {
      event.preventDefault();
      event.stopPropagation();
      if (currentIndex === total - 1 && taskIsComplete(currentIndex)) {
        location.hash = "#1";
        location.reload();
        return;
      }
      next();
    });
  }

  function updateChrome() {
    const slide = slides[currentIndex];
    currentNumber.textContent = String(currentIndex + 1).padStart(2, "0");
    totalNumber.textContent = String(total).padStart(2, "0");
    sectionPill.textContent = slide.dataset.section || "";
    progressBar.style.width = `${((currentIndex + 1) / total) * 100}%`;
    stage.classList.toggle("is-cover", currentIndex === 0);
    document.title = `${slide.dataset.title}｜阿狸老师的作文小教室`;
    const targetHash = `#${currentIndex + 1}`;
    if (location.hash !== targetHash) {
      try {
        history.replaceState(null, "", targetHash);
      } catch {
        // HTMLPreview 等公网代理会把文档基地址保留为 raw.githubusercontent.com，
        // 此时 replaceState 会触发跨源异常。退回普通 hash 更新，不能让导航栏停止刷新。
        location.hash = targetHash;
      }
    }
    notesTitle.textContent = slide.dataset.title || "讲者备注";
    notesTime.textContent = slide.dataset.time || "";
    const note = slide.querySelector("template.speaker-note");
    notesContent.textContent = note ? note.content.textContent.trim() : "此页暂无讲者备注。";
    outlineList.querySelectorAll(".outline-item").forEach((item, index) => item.classList.toggle("is-current", index === currentIndex));
    updateCoach();
  }

  function showSlide(nextIndex, direction = 1, initial = false) {
    const bounded = Math.max(0, Math.min(total - 1, nextIndex));
    if (bounded === currentIndex && !initial) return;
    const oldSlide = slides[currentIndex];
    const nextSlide = slides[bounded];
    if (transitionTimer !== null) {
      clearTimeout(transitionTimer);
      transitionTimer = null;
    }
    transitioning = true;
    slides.forEach(slide => slide.classList.remove("is-active", "is-leaving-left", "is-leaving-right"));
    if (!initial && oldSlide !== nextSlide) oldSlide.classList.add(direction > 0 ? "is-leaving-left" : "is-leaving-right");
    currentIndex = bounded;
    nextSlide.classList.add("is-active");
    updateChrome();
    if (!initial) playSound(direction > 0 ? "page" : "back");
    transitionTimer = setTimeout(() => {
      slides.forEach((slide, index) => {
        slide.classList.toggle("is-active", index === currentIndex);
        slide.classList.remove("is-leaving-left", "is-leaving-right");
      });
      transitioning = false;
      transitionTimer = null;
    }, initial ? 20 : 430);
  }

  function next() {
    if (!taskIsComplete(currentIndex)) {
      coachPanel.classList.remove("needs-attention");
      void coachPanel.offsetWidth;
      coachPanel.classList.add("needs-attention");
      updateTargetHints(currentIndex);
      showToast("先完成本页点击任务；赶时间可点“跳过”");
      playSound("back");
      return;
    }
    if (currentIndex < total - 1) showSlide(currentIndex + 1, 1);
  }

  function previous() {
    if (currentIndex > 0) showSlide(currentIndex - 1, -1);
  }

  async function toggleFullscreen() {
    try {
      if (!document.fullscreenElement) await document.documentElement.requestFullscreen();
      else await document.exitFullscreen();
      playSound("click");
    } catch {
      showToast("浏览器未允许全屏，可用浏览器菜单进入全屏");
    }
  }

  function toggleSound() {
    if (soundEnabled) {
      playSound("off", true);
      soundEnabled = false;
      soundButton.classList.remove("sound-on");
      soundButton.setAttribute("aria-pressed", "false");
      showToast("提示音已关闭");
    } else {
      soundEnabled = true;
      soundButton.classList.add("sound-on");
      soundButton.setAttribute("aria-pressed", "true");
      playSound("success");
      showToast("提示音已开启");
    }
  }

  function buildOutline() {
    const fragment = document.createDocumentFragment();
    slides.forEach((slide, index) => {
      const button = document.createElement("button");
      button.type = "button";
      button.className = "outline-item";
      button.innerHTML = `<b>${String(index + 1).padStart(2, "0")}</b><span>${slide.dataset.title}</span><small>${slide.dataset.time || ""}</small>`;
      button.addEventListener("click", () => {
        showSlide(index, index >= currentIndex ? 1 : -1);
        closeDrawers();
      });
      fragment.appendChild(button);
    });
    outlineList.appendChild(fragment);
  }

  const scoreData = {
    small: [
      ["7–10分", "第三档", "全部题干信息 · 类型明确 · 表达恰当"],
      ["4–6分", "第二档", "大部分信息 · 类型基本明确 · 表达通顺"],
      ["0–3分", "第一档", "信息很少 · 类型不明 · 表达不通顺或少于80字"]
    ],
    large: [
      ["26–30分", "第六档", "全部要点 · 准确流畅 · 形式丰富"],
      ["20–25分", "第五档", "全部要点 · 表达恰当"],
      ["15–19分", "第四档", "大部分要点 · 表达通顺"],
      ["10–14分", "第三档", "一部分要点 · 基本通顺"],
      ["5–9分", "第二档", "少部分要点 · 欠通顺"],
      ["0–4分", "第一档", "要点很少 · 不通顺或少于100字"]
    ]
  };

  function renderScores(type) {
    const bands = document.getElementById("scoreBands");
    bands.dataset.mode = type;
    bands.innerHTML = scoreData[type].map(row => `<article class="score-band"><b>${row[0]}</b><span><strong>${row[1]}</strong>${row[2]}</span></article>`).join("");
  }

  function bindScores() {
    document.querySelectorAll(".score-tab").forEach(button => button.addEventListener("click", () => {
      document.querySelectorAll(".score-tab").forEach(item => item.classList.remove("is-active"));
      button.classList.add("is-active");
      renderScores(button.dataset.score);
    }));
    renderScores("small");
  }

  const ruleData = {
    indent: ["只空一格", "常规作文每段开头空一格；电子邮件通常顶格。"],
    punctuation: ["各占一格", "“、”与“。”各占一格，标点不能单独出现在行首。"],
    smallkana: ["小也占一格", "拗音、促音写得小，但仍然单独占一个格子。"],
    number: ["一格最多两个", "横写阿拉伯数字时，一个格子最多放两个数字。"],
    quote: ["成对出现", "「」与『』要成对，内外两层不要混用。"]
  };

  function bindRules() {
    const display = document.getElementById("ruleDisplay");
    const paper = document.getElementById("miniGenko");
    document.querySelectorAll(".rule-tab").forEach(button => button.addEventListener("click", () => {
      document.querySelectorAll(".rule-tab").forEach(item => item.classList.remove("is-active"));
      button.classList.add("is-active");
      const [title, copy] = ruleData[button.dataset.rule];
      display.innerHTML = `<b>${title}</b><p>${copy}</p>`;
      paper.dataset.highlight = button.dataset.rule;
    }));
  }

  function bindErrors() {
    document.querySelectorAll(".error-item").forEach(button => {
      const answer = button.querySelector("ins");
      answer.dataset.why = button.dataset.why;
      button.addEventListener("click", () => {
        button.classList.add("is-open");
        answer.textContent = `→ ${button.dataset.fix}`;
      });
    });
  }

  function bindChecklist() {
    const buttons = Array.from(document.querySelectorAll("#checkGrid button"));
    const count = document.getElementById("checkCount");
    const title = document.getElementById("checkTitle");
    const copy = document.getElementById("checkText");
    const meter = document.querySelector(".paw-meter");
    function update() {
      const done = buttons.filter(button => button.classList.contains("is-checked")).length;
      count.textContent = String(done);
      meter.style.borderColor = done === 6 ? "#ef8d4d" : "#fae2e0";
      title.textContent = done === 6 ? "检查完成，可以誊清" : done === 0 ? "从“审题”开始" : `已完成 ${done} 组`;
      copy.textContent = done === 6 ? "现在只做小修正，保持卷面整洁。" : `还剩 ${6 - done} 组，继续按顺序检查。`;
    }
    buttons.forEach(button => button.addEventListener("click", () => {
      button.classList.add("is-checked");
      update();
    }));
    document.getElementById("resetCheck").addEventListener("click", () => {
      buttons.forEach(button => button.classList.remove("is-checked"));
      update();
      playSound("back");
    });
  }

  function bindControls() {
    document.getElementById("prevSlide").addEventListener("click", previous);
    document.getElementById("nextSlide").addEventListener("click", next);
    document.querySelectorAll("[data-next]").forEach(button => button.addEventListener("click", next));
    document.getElementById("openOutline").addEventListener("click", () => {
      const open = !outlineDrawer.classList.contains("is-open");
      setDrawer(notesDrawer, false); setDrawer(outlineDrawer, open); playSound("click");
    });
    document.getElementById("toggleNotes").addEventListener("click", () => {
      const open = !notesDrawer.classList.contains("is-open");
      setDrawer(outlineDrawer, false); setDrawer(notesDrawer, open); playSound("click");
    });
    soundButton.addEventListener("click", toggleSound);
    document.getElementById("toggleFullscreen").addEventListener("click", toggleFullscreen);
    document.querySelectorAll("[data-close='outline']").forEach(button => button.addEventListener("click", () => setDrawer(outlineDrawer, false)));
    document.querySelectorAll("[data-close='notes']").forEach(button => button.addEventListener("click", () => setDrawer(notesDrawer, false)));

    document.addEventListener("keydown", event => {
      if (["ArrowRight", " ", "PageDown"].includes(event.key)) { event.preventDefault(); next(); }
      else if (["ArrowLeft", "PageUp"].includes(event.key)) { event.preventDefault(); previous(); }
      else if (event.key === "Home") { event.preventDefault(); showSlide(0, -1); }
      else if (event.key === "End") { event.preventDefault(); showSlide(total - 1, 1); }
      else if (event.key.toLowerCase() === "o") setDrawer(outlineDrawer, !outlineDrawer.classList.contains("is-open"));
      else if (event.key.toLowerCase() === "n") setDrawer(notesDrawer, !notesDrawer.classList.contains("is-open"));
      else if (event.key.toLowerCase() === "f") toggleFullscreen();
      else if (event.key.toLowerCase() === "m") toggleSound();
      else if (event.key === "Escape") closeDrawers();
    });

    stage.addEventListener("touchstart", event => { touchStartX = event.changedTouches[0].clientX; }, { passive: true });
    stage.addEventListener("touchend", event => {
      if (touchStartX === null) return;
      const delta = event.changedTouches[0].clientX - touchStartX;
      touchStartX = null;
      if (Math.abs(delta) < 48) return;
      if (delta < 0) next(); else previous();
    }, { passive: true });
  }

  window.addEventListener("resize", fitStage);
  window.addEventListener("hashchange", () => {
    const index = Math.max(0, Math.min(total - 1, Number(location.hash.slice(1)) - 1 || 0));
    if (index !== currentIndex) showSlide(index, index > currentIndex ? 1 : -1);
  });

  fitStage();
  buildOutline();
  bindGuidedTasks();
  bindControls();
  bindScores();
  bindRules();
  bindErrors();
  bindChecklist();
  showSlide(currentIndex, 1, true);
})();
