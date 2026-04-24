import { invoke } from "@tauri-apps/api/core";

type Route = "home" | "customs" | "processor" | "settings";

type GenerateResult = {
  boeFiles: string[];
  fitonFiles: string[];
  invCount: number;
  lineCount: number;
  warnings: string[];
};

type ProcessResult = {
  invFiles: number;
  plFiles: number;
  invRows: number;
  plRows: number;
  skippedFiles: number;
  outputPath: string;
};

type AppState = {
  displayName: string;
  route: Route;
  // customs generator
  customsInputDir: string;
  mawbPath: string;
  customsOutputDir: string;
  lastGenResult: GenerateResult | null;
  // legacy processor
  inputDir: string;
  outputPath: string;
  lastResult: ProcessResult | null;
  status: string;
};

const STORAGE_KEY = "custom_app_state_v3";
const routeLabels: Record<Route, string> = {
  home: "主页",
  customs: "报关文件生成",
  processor: "旧版处理",
  settings: "设置",
};

const defaultState: AppState = {
  displayName: "",
  route: "home",
  customsInputDir: "",
  mawbPath: "",
  customsOutputDir: "",
  lastGenResult: null,
  inputDir: "",
  outputPath: "",
  lastResult: null,
  status: "",
};

let state: AppState = loadState();
let isProcessing = false;

function loadState(): AppState {
  const raw = localStorage.getItem(STORAGE_KEY);
  if (!raw) return { ...defaultState };
  try {
    const p = JSON.parse(raw) as Partial<AppState>;
    const route: Route =
      p.route === "home" || p.route === "customs" || p.route === "processor" || p.route === "settings"
        ? p.route : "home";
    return {
      displayName: typeof p.displayName === "string" ? p.displayName : "",
      route,
      customsInputDir: typeof p.customsInputDir === "string" ? p.customsInputDir : "",
      mawbPath: typeof p.mawbPath === "string" ? p.mawbPath : "",
      customsOutputDir: typeof p.customsOutputDir === "string" ? p.customsOutputDir : "",
      lastGenResult: p.lastGenResult ?? null,
      inputDir: typeof p.inputDir === "string" ? p.inputDir : "",
      outputPath: typeof p.outputPath === "string" ? p.outputPath : "",
      lastResult: p.lastResult ?? null,
      status: typeof p.status === "string" ? p.status : "",
    };
  } catch {
    return { ...defaultState };
  }
}

function saveState(): void {
  localStorage.setItem(STORAGE_KEY, JSON.stringify(state));
}

function setRoute(route: Route): void {
  state.route = route;
  saveState();
  render();
}

function setWindowTitle(): void {
  document.title = `Custom App | ${routeLabels[state.route]}`;
}

// ── Views ────────────────────────────────────────────────────────────────────

function viewHome(): string {
  const greeting = state.displayName ? `你好，${escapeHtml(state.displayName)}。` : "你好，欢迎使用 Custom App。";
  return `
    <section class="panel">
      <h2>欢迎</h2>
      <p class="muted">${greeting}</p>
      <label class="field-label" for="display-name">昵称</label>
      <input id="display-name" class="text-input" type="text" placeholder="输入你的昵称" value="${escapeAttr(state.displayName)}" />
      <button id="save-name" class="btn btn-primary" type="button">保存昵称</button>
    </section>
  `;
}

function viewCustoms(): string {
  const r = state.lastGenResult;
  const resultBlock = r ? `
    <div class="status-card">
      <div>发票数：${r.invCount}，行项目：${r.lineCount}</div>
      ${r.boeFiles.map(f => `<div>BOE: <span class="path">${escapeHtml(f)}</span></div>`).join("")}
      ${r.fitonFiles.map(f => `<div>Fiton: <span class="path">${escapeHtml(f)}</span></div>`).join("")}
      ${r.warnings.length ? `<div class="warn">⚠ ${r.warnings.map(escapeHtml).join("<br>⚠ ")}</div>` : ""}
    </div>
  ` : "";

  return `
    <section class="panel">
      <h2>报关文件生成</h2>
      <p class="muted">从发票和装箱单自动生成 BOE Draft Form 和 Fiton 报关文件。</p>
      <div class="form-grid">
        <div>
          <label class="field-label" for="customs-input-dir">发票目录（含 INV 和 packing list 文件）</label>
          <input id="customs-input-dir" class="text-input" type="text"
            placeholder="例如 /home/zxh/data/shipment"
            value="${escapeAttr(state.customsInputDir)}" />
        </div>
        <div>
          <label class="field-label" for="mawb-path">MAWB PDF 路径</label>
          <input id="mawb-path" class="text-input" type="text"
            placeholder="例如 /home/zxh/data/784-83551123 MAWB.pdf"
            value="${escapeAttr(state.mawbPath)}" />
        </div>
        <div>
          <label class="field-label" for="customs-output-dir">输出目录</label>
          <input id="customs-output-dir" class="text-input" type="text"
            placeholder="例如 /home/zxh/data/output"
            value="${escapeAttr(state.customsOutputDir)}" />
        </div>
      </div>
      <button id="run-customs" class="btn btn-primary" type="button" ${isProcessing ? "disabled" : ""}>
        ${isProcessing ? "生成中..." : "生成报关文件"}
      </button>
      <p class="footnote">${escapeHtml(state.status)}</p>
      ${resultBlock}
    </section>
  `;
}

function viewProcessor(): string {
  const r = state.lastResult;
  const resultBlock = r ? `
    <div class="status-card">
      <div>INV 文件数：${r.invFiles}，Packing List：${r.plFiles}</div>
      <div>Sheet1 行数：${r.invRows}，Sheet2 行数：${r.plRows}，跳过：${r.skippedFiles}</div>
      <div>输出：<span class="path">${escapeHtml(r.outputPath)}</span></div>
    </div>
  ` : "";

  return `
    <section class="panel">
      <h2>旧版文件处理</h2>
      <p class="muted">生成 output.xlsx（Sheet1 发票数据 + Sheet2 装箱单数据）。</p>
      <div class="form-grid">
        <div>
          <label class="field-label" for="input-dir">输入目录</label>
          <input id="input-dir" class="text-input" type="text"
            placeholder="例如 /home/zxh/data"
            value="${escapeAttr(state.inputDir)}" />
        </div>
        <div>
          <label class="field-label" for="output-path">输出文件</label>
          <input id="output-path" class="text-input" type="text"
            placeholder="例如 /home/zxh/data/output.xlsx"
            value="${escapeAttr(state.outputPath)}" />
        </div>
      </div>
      <button id="run-process" class="btn btn-primary" type="button" ${isProcessing ? "disabled" : ""}>
        ${isProcessing ? "处理中..." : "开始处理"}
      </button>
      <p class="footnote">${escapeHtml(state.status)}</p>
      ${resultBlock}
    </section>
  `;
}

function viewSettings(): string {
  return `
    <section class="panel">
      <h2>设置</h2>
      <p class="muted">清空本地缓存状态。</p>
      <button id="clear-data" class="btn btn-danger" type="button">清空本地数据</button>
    </section>
  `;
}

function viewContent(): string {
  if (state.route === "home") return viewHome();
  if (state.route === "customs") return viewCustoms();
  if (state.route === "processor") return viewProcessor();
  return viewSettings();
}

// ── Render ───────────────────────────────────────────────────────────────────

function render(): void {
  const app = document.querySelector<HTMLDivElement>("#app");
  if (!app) return;
  setWindowTitle();
  app.innerHTML = `
    <div class="app-shell">
      <header class="topbar">
        <div class="brand">Custom App</div>
        <nav class="menu" aria-label="主菜单">
          ${(["home", "customs", "processor", "settings"] as Route[]).map(r => `
            <button class="menu-item ${state.route === r ? "active" : ""}" data-route="${r}" type="button">
              ${routeLabels[r]}
            </button>
          `).join("")}
        </nav>
      </header>
      <main class="content">${viewContent()}</main>
    </div>
  `;
  bindEvents();
}

// ── Events ───────────────────────────────────────────────────────────────────

function bindEvents(): void {
  document.querySelectorAll<HTMLButtonElement>(".menu-item").forEach(btn => {
    btn.addEventListener("click", () => {
      const route = btn.dataset.route as Route | undefined;
      if (route) setRoute(route);
    });
  });

  // Home
  const saveNameBtn = document.querySelector<HTMLButtonElement>("#save-name");
  const nameInput = document.querySelector<HTMLInputElement>("#display-name");
  if (saveNameBtn && nameInput) {
    saveNameBtn.addEventListener("click", () => {
      state.displayName = nameInput.value.trim();
      state.status = "昵称已保存";
      saveState();
      render();
    });
  }

  // Customs inputs
  bindInput("#customs-input-dir", v => { state.customsInputDir = v; });
  bindInput("#mawb-path", v => { state.mawbPath = v; });
  bindInput("#customs-output-dir", v => { state.customsOutputDir = v; });

  const runCustomsBtn = document.querySelector<HTMLButtonElement>("#run-customs");
  if (runCustomsBtn) runCustomsBtn.addEventListener("click", runCustoms);

  // Legacy processor inputs
  bindInput("#input-dir", v => { state.inputDir = v; });
  bindInput("#output-path", v => { state.outputPath = v; });

  const runProcessBtn = document.querySelector<HTMLButtonElement>("#run-process");
  if (runProcessBtn) runProcessBtn.addEventListener("click", runProcess);

  // Settings
  const clearBtn = document.querySelector<HTMLButtonElement>("#clear-data");
  if (clearBtn) {
    clearBtn.addEventListener("click", () => {
      localStorage.removeItem(STORAGE_KEY);
      state = { ...defaultState };
      render();
    });
  }
}

function bindInput(selector: string, setter: (v: string) => void): void {
  const el = document.querySelector<HTMLInputElement>(selector);
  if (el) {
    el.addEventListener("input", () => {
      setter(el.value.trim());
      saveState();
    });
  }
}

// ── Actions ──────────────────────────────────────────────────────────────────

async function runCustoms(): Promise<void> {
  if (!state.customsInputDir) { state.status = "请填写发票目录"; render(); return; }
  if (!state.customsOutputDir) { state.status = "请填写输出目录"; render(); return; }

  isProcessing = true;
  state.status = "正在生成报关文件，请稍候...";
  state.lastGenResult = null;
  saveState();
  render();

  try {
    const result = await invoke<GenerateResult>("generate_customs_docs", {
      inputDir: state.customsInputDir,
      mawbPath: state.mawbPath,
      outputDir: state.customsOutputDir,
    });
    state.lastGenResult = result;
    state.status = `完成：生成 ${result.boeFiles.length} 份 BOE + ${result.fitonFiles.length} 份 Fiton`;
  } catch (error) {
    state.status = `生成失败：${String(error)}`;
  } finally {
    isProcessing = false;
    saveState();
    render();
  }
}

async function runProcess(): Promise<void> {
  if (!state.inputDir) { state.status = "请先填写输入目录"; render(); return; }
  if (!state.outputPath) { state.status = "请先填写输出文件路径"; render(); return; }

  isProcessing = true;
  state.status = "正在处理文件，请稍候...";
  saveState();
  render();

  try {
    const result = await invoke<ProcessResult>("process_excel_files", {
      inputDir: state.inputDir,
      outputPath: state.outputPath,
    });
    state.lastResult = result;
    state.status = "处理完成";
  } catch (error) {
    state.status = `处理失败：${String(error)}`;
  } finally {
    isProcessing = false;
    saveState();
    render();
  }
}

// ── Utils ────────────────────────────────────────────────────────────────────

function escapeHtml(value: string): string {
  return value.replace(/&/g, "&amp;").replace(/</g, "&lt;").replace(/>/g, "&gt;")
    .replace(/"/g, "&quot;").replace(/'/g, "&#39;");
}

function escapeAttr(value: string): string {
  return escapeHtml(value);
}

window.addEventListener("DOMContentLoaded", () => { render(); });
