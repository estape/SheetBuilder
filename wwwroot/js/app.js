/**
 * App.js - Controlador do Cliente SheetBuilder
 * Conecta eventos de interface à ponte IPC do Photino.NET e orquestra a experiência do usuário.
 */

class AppController {
  constructor() {
    this.view = new View();
    this.pendingRequests = new Map();
    this.currentExcelPath = null;

    this.initPhotinoBridge();
    this.bindUserActions();

    // Consulta o ambiente Python ao iniciar
    this.requestEnvironment();
  }

  // --- Ponte IPC Photino.NET ---

  initPhotinoBridge() {
    if (window.external && typeof window.external.receiveMessage === "function") {
      window.external.receiveMessage((rawMessage) => {
        try {
          const data = typeof rawMessage === "string" ? JSON.parse(rawMessage) : rawMessage;
          this.handleIncomingMessage(data);
        } catch (e) {
          console.error("Erro ao interpretar mensagem do backend .NET:", e, rawMessage);
        }
      });
    } else {
      console.warn("Ponte Photino não identificada neste ambiente (ex: navegador externo).");
    }
  }

  sendToDotNet(action, payload = {}) {
    return new Promise((resolve, reject) => {
      const id = "msg_" + Math.random().toString(36).substring(2, 11);
      const message = JSON.stringify({ action, id, payload });

      this.pendingRequests.set(id, { resolve, reject, action });

      if (window.external && typeof window.external.sendMessage === "function") {
        window.external.sendMessage(message);
      } else {
        console.warn("Simulação de resposta para:", action);
        this.pendingRequests.delete(id);
        reject(new Error("Ponte nativa Photino não conectada"));
      }

      // Tempo limite de segurança de 2 minutos para processamentos analíticos extensos
      setTimeout(() => {
        if (this.pendingRequests.has(id)) {
          this.pendingRequests.delete(id);
          reject(new Error(`Tempo limite excedido para a ação: ${action}`));
        }
      }, 120000);
    });
  }

  handleIncomingMessage(data) {
    // Eventos contínuos (ex: logs e progresso de processamento)
    if (data.event === "log" && data.payload) {
      this.view.appendLog(data.payload.message, data.payload.isError);
      return;
    }

    // Correlação com requisição pendente
    if (data.id && this.pendingRequests.has(data.id)) {
      const { resolve, reject } = this.pendingRequests.get(data.id);
      this.pendingRequests.delete(data.id);

      if (data.success) {
        resolve(data.payload);
      } else {
        reject(data.payload || new Error("Falha na execução da ação solicitada"));
      }
    }
  }

  // --- Ações do Usuário ---

  bindUserActions() {
    // 1. Gabaritos Financeiros
    this.view.presetSelect.addEventListener("change", (e) => {
      const key = e.target.value;
      if (key && PRESETS[key]) {
        this.view.codeInput.value = PRESETS[key];
        this.view.updateLineNumbers();
        this.view.showToast("Gabarito financeiro carregado com sucesso!");
      }
    });

    this.view.quickPresetBtn.addEventListener("click", () => {
      this.view.presetSelect.value = "dre_gerencial";
      this.view.codeInput.value = PRESETS.dre_gerencial;
      this.view.updateLineNumbers();
      this.view.showToast("Gabarito de DRE Gerencial carregado!");
    });

    // 2. Colar e Limpar Código
    this.view.pasteBtn.addEventListener("click", async () => {
      try {
        const text = await navigator.clipboard.readText();
        if (text) {
          this.view.codeInput.value = text;
          this.view.updateLineNumbers();
          this.view.showToast("Código colado da área de transferência!");
        }
      } catch (err) {
        this.view.showToast("Permissão de área de transferência não disponível. Utilize o atalho Ctrl+V.", true);
      }
    });

    this.view.clearBtn.addEventListener("click", () => {
      this.view.codeInput.value = "";
      this.view.updateLineNumbers();
    });

    // 3. Gerar Planilha Executiva
    this.view.executeBtn.addEventListener("click", () => this.executeScript());

    // 4. Salvar Pasta de Trabalho (Diálogo nativo do Windows 11 com padrão em Documentos)
    this.view.saveExcelBtn.addEventListener("click", () => this.saveExcelDialog());

    // 5. Abrir no Microsoft Excel
    this.view.openExcelBtn.addEventListener("click", () => this.openInExcel());

    // 6. Abrir Pasta no Explorador de Arquivos
    this.view.openFolderBtn.addEventListener("click", () => this.openFolder());
  }

  async requestEnvironment() {
    try {
      const env = await this.sendToDotNet("get_env");
      if (env && env.isInstalled) {
        this.view.envBadge.className = "env-badge online";
        const pkgDetails = [];
        if (env.hasOpenPyXL) pkgDetails.push("openpyxl");
        if (env.hasPandas) pkgDetails.push("pandas");
        const details = pkgDetails.length ? `(${pkgDetails.join(", ")})` : "";
        this.view.envText.textContent = `Motor Python ${env.versionString.split(" ")[0]} Ativo ${details}`;
      } else {
        this.view.envBadge.className = "env-badge offline";
        this.view.envText.textContent = "Python não localizado no sistema";
      }
    } catch (e) {
      this.view.envBadge.className = "env-badge offline";
      this.view.envText.textContent = "Erro na verificação do ambiente";
    }
  }

  async executeScript() {
    const code = this.view.codeInput.value.trim();
    if (!code) {
      this.view.showToast("Por favor, cole o código do Google AI Studio antes de gerar.", true);
      this.view.codeInput.focus();
      return;
    }

    this.view.setExecutingState(true);
    this.view.switchTab("console");
    this.view.terminalOutput.textContent = "";
    this.view.appendLog(`[INÍCIO] Processando modelo analítico às ${new Date().toLocaleTimeString()}...`, false);

    const envVars = {};
    const geminiKey = this.view.geminiKeyInput.value.trim();
    if (geminiKey) {
      envVars["GEMINI_API_KEY"] = geminiKey;
    }

    try {
      const result = await this.sendToDotNet("execute_script", {
        code,
        timeout: 90,
        envVars
      });

      this.view.setExecutingState(false);

      if (result.isSuccess && result.generatedExcelPath) {
        this.currentExcelPath = result.generatedExcelPath;
        this.view.updateFileActions(this.currentExcelPath);
        this.view.execStatusBadge.textContent = "Planilha Gerada!";
        this.view.execStatusBadge.className = "status-badge";
        this.view.showToast("Planilha financeira gerada com sucesso!");

        // Renderizar conferência
        if (result.sheetData) {
          this.view.renderPreview(result.sheetData, (sheetName) => this.switchSheet(sheetName));
        }
      } else if (result.isSuccess) {
        this.view.execStatusBadge.textContent = "Concluído (Sem arquivo Excel)";
        this.view.showToast("Processamento finalizado, mas nenhum arquivo .xlsx foi detectado.", true);
      } else {
        this.view.execStatusBadge.textContent = "Requer Revisão de Premissas";
        this.view.showToast(`Falha no processamento: ${result.errorMessage || "Consulte o diário de bordo"}`, true);
      }
    } catch (err) {
      this.view.setExecutingState(false);
      this.view.execStatusBadge.textContent = "Erro de Processamento";
      this.view.appendLog(`[ERRO] ${err.message}`, true);
      this.view.showToast(err.message, true);
    }
  }

  async switchSheet(sheetName) {
    if (!this.currentExcelPath) return;
    try {
      const preview = await this.sendToDotNet("get_sheet_preview", {
        filePath: this.currentExcelPath,
        sheetName
      });
      if (preview) {
        this.view.renderPreview(preview, (sName) => this.switchSheet(sName));
      }
    } catch (err) {
      this.view.showToast(`Erro ao alternar planilha: ${err.message}`, true);
    }
  }

  async saveExcelDialog() {
    if (!this.currentExcelPath) {
      this.view.showToast("Nenhuma planilha disponível para salvar no momento", true);
      return;
    }

    try {
      const res = await this.sendToDotNet("save_excel_dialog", {
        filePath: this.currentExcelPath
      });

      if (res && res.savedPath) {
        this.currentExcelPath = res.savedPath;
        this.view.updateFileActions(this.currentExcelPath);
        this.view.showToast(`Pasta de trabalho salva com sucesso em: ${res.savedPath}`);
      } else if (res && res.cancelled) {
        // Cancelado pelo usuário
      }
    } catch (err) {
      this.view.showToast(`Falha ao salvar arquivo: ${err.message}`, true);
    }
  }

  async openInExcel() {
    if (!this.currentExcelPath) return;

    try {
      const res = await this.sendToDotNet("open_in_excel", {
        filePath: this.currentExcelPath
      });
      if (res && res.opened) {
        this.view.showToast("Abrindo a pasta de trabalho no Microsoft Excel...");
      } else {
        this.view.showToast("Não foi possível iniciar o Microsoft Excel no sistema", true);
      }
    } catch (err) {
      this.view.showToast(err.message, true);
    }
  }

  async openFolder() {
    if (!this.currentExcelPath) return;

    try {
      await this.sendToDotNet("open_folder", {
        filePath: this.currentExcelPath
      });
    } catch (err) {
      this.view.showToast(err.message, true);
    }
  }
}

// Inicializar na carga do DOM
window.addEventListener("DOMContentLoaded", () => {
  window.sheetBuilderApp = new AppController();
});
