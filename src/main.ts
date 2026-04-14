import { invoke } from "@tauri-apps/api/core";

type Route = "home" | "processor" | "settings";

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
  inputDir: string;
  outputPath: string;
  status: string;
  lastResult: ProcessResult | null;
};

const STORAGE_KEY = "custom_app_state_v2";
const routeLabels: Record<Route, string> = {
  home: "主页",
  processor: "文件处理",
  settings: "设置",
};

const defaultState: AppState = {
  displayName: "",
  route: "home",
  inputDir: "",
  outputPath: "",
  status: "",
  lastResult: null,
};

let state: AppState = loadState();
let isProcessing = false;

function loadState(): AppState {
  const raw = localStorage.getItem(STORAGE_KEY);
  if (!raw) {
    return { ...defaultState };
  }

  try {
    const parsed = JSON.parse(raw) as Partial<AppState>;
    const route: Route =
      parsed.route === "home" || parsed.route === "processor" || parsed.route === "settings"
        ? parsed.route
        : "home";

    return {
      displayName: typeof parsed.displayName === "string" ? parsed.displayName : "",
      route,
      inputDir: typeof parsed.inputDir === "string" ? parsed.inputDir : "",
      outputPath: typeof parsed.outputPath === "string" ? parsed.outputPath : "",
      status: typeof parsed.status === "string" ? parsed.status : "",
      lastResult:
        parsed.lastResult && typeof parsed.lastResult === "object"
          ? (parsed.lastResult as ProcessResult)
          : null,
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

function viewProcessor(): string {
  const resultBlock = state.lastResult
    ? `
      <div class="status-card">
        <div>INV 文件数：${state.lastResult.invFiles}</div>
        <div>Packing List 文件数：${state.lastResult.plFiles}</div>
        <div>Sheet1 行数：${state.lastResult.invRows}</div>
        <div>Sheet2 行数：${state.lastResult.plRows}</div>
        <div>跳过文件：${state.lastResult.skippedFiles}</div>
        <div>输出路径：${escapeHtml(state.lastResult.outputPath)}</div>
      </div>
    `
    : "";

  return `
    <section class="panel">
      <h2>文件处理</h2>
      <p class="muted">按你提供的 extract.py 逻辑处理目录中的 Excel，并生成 output 文件。</p>
      <div class="processor-grid">
        <div>
          <label class="field-label" for="input-dir">输入目录</label>
          <input id="input-dir" class="text-input" type="text" placeholder="例如 /home/zxh/data" value="${escapeAttr(state.inputDir)}" />
        </div>
        <div>
          <label class="field-label" for="output-path">输出文件</label>
          <input id="output-path" class="text-input" type="text" placeholder="例如 /home/zxh/data/output.xlsx" value="${escapeAttr(state.outputPath)}" />
        </div>
      </div>
      <button id="run-process" class="btn btn-primary" type="button" ${isProcessing ? "disabled" : ""}>${isProcessing ? "处理中..." : "开始处理"}</button>
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
  if (state.route === "home") {
    return viewHome();
  }
  if (state.route === "processor") {
    return viewProcessor();
  }
  return viewSettings();
}

function render(): void {
  const app = document.querySelector<HTMLDivElement>("#app");
  if (!app) {
    return;
  }

  setWindowTitle();

  app.innerHTML = `
    <div class="app-shell">
      <header class="topbar">
        <div class="brand">Custom App</div>
        <nav class="menu" aria-label="主菜单">
          <button class="menu-item ${state.route === "home" ? "active" : ""}" data-route="home" type="button">主页</button>
          <button class="menu-item ${state.route === "processor" ? "active" : ""}" data-route="processor" type="button">文件处理</button>
          <button class="menu-item ${state.route === "settings" ? "active" : ""}" data-route="settings" type="button">设置</button>
        </nav>
      </header>
      <main class="content">${viewContent()}</main>
    </div>
  `;

  bindEvents();
}

function bindEvents(): void {
  document.querySelectorAll<HTMLButtonElement>(".menu-item").forEach((button) => {
    button.addEventListener("click", () => {
      const route = button.dataset.route as Route | undefined;
      if (route) {
        setRoute(route);
      }
    });
  });

  const saveNameButton = document.querySelector<HTMLButtonElement>("#save-name");
  const nameInput = document.querySelector<HTMLInputElement>("#display-name");
  if (saveNameButton && nameInput) {
    saveNameButton.addEventListener("click", () => {
      state.displayName = nameInput.value.trim();
      state.status = "昵称已保存";
      saveState();
      render();
    });
  }

  const inputDir = document.querySelector<HTMLInputElement>("#input-dir");
  if (inputDir) {
    inputDir.addEventListener("input", () => {
      state.inputDir = inputDir.value.trim();
      saveState();
    });
  }

  const outputPath = document.querySelector<HTMLInputElement>("#output-path");
  if (outputPath) {
    outputPath.addEventListener("input", () => {
      state.outputPath = outputPath.value.trim();
      saveState();
    });
  }

  const runButton = document.querySelector<HTMLButtonElement>("#run-process");
  if (runButton) {
    runButton.addEventListener("click", runProcess);
  }

  const clearDataButton = document.querySelector<HTMLButtonElement>("#clear-data");
  if (clearDataButton) {
    clearDataButton.addEventListener("click", () => {
      localStorage.removeItem(STORAGE_KEY);
      state = { ...defaultState };
      render();
    });
  }
}

async function runProcess(): Promise<void> {
  if (!state.inputDir) {
    state.status = "请先填写输入目录";
    render();
    return;
  }
  if (!state.outputPath) {
    state.status = "请先填写输出文件路径";
    render();
    return;
  }

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

function escapeHtml(value: string): string {
  return value
    .split("&")
    .join("&amp;")
    .split("<")
    .join("&lt;")
    .split(">")
    .join("&gt;")
    .split('"')
    .join("&quot;")
    .split("'")
    .join("&#39;");
}

function escapeAttr(value: string): string {
  return escapeHtml(value);
}

window.addEventListener("DOMContentLoaded", () => {
  render();
});
