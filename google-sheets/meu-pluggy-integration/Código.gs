// ==========================================
// 1. CONFIGURAÇÕES GLOBAIS E CREDENCIAIS
// ==========================================
const PLUGGY_BASE_URL = "https://api.pluggy.ai";
const DAYS_TO_SYNC = 30;

function setupCredentials() {
  PropertiesService.getScriptProperties().setProperties({
    "PLUGGY_CLIENT_ID": "SEU_CLIENT_ID_AQUI",
    "PLUGGY_CLIENT_SECRET": "SEU_CLIENT_SECRET_AQUI"
  });
}

function setupManualItemsIds() {
  const ids = [
    "550e8400-e29b-41d4-a716-446655440000",
    "6ba7b810-9dad-41d1-80b4-00c04fd430c8"
  ];

  PropertiesService.getScriptProperties().setProperty(
    "MANUAL_ITEMS_IDS",
    JSON.stringify(ids)
  );
}

function setupSensitiveGlobalVariables() {
  setupCredentials();
  setupManualItemsIds();
}

function cleanUp() {
  const props = PropertiesService.getScriptProperties();
  props.deleteProperty("PLUGGY_CLIENT_ID");
  props.deleteProperty("PLUGGY_CLIENT_SECRET");
  props.deleteProperty("MANUAL_ITEMS_IDS");
}

// ==========================================
// 2. CLIENTE HTTP COM RETRY & BACKOFF
// ==========================================
function fetchWithRetry(url, options, maxRetries = 3, ignoreNotFound = false) {
  let code;
  for (let i = 0; i < maxRetries; i++) {
    try {
      const requestOptions = { ...options, muteHttpExceptions: true };
      const response = UrlFetchApp.fetch(url, requestOptions);
      code = response.getResponseCode();
      const text = response.getContentText();

      if (code >= 200 && code < 300) {
        return JSON.parse(text);
      }

      // Se um endpoint opcional (ex: loans/investments) não existir ou for vazio para o item
      if (ignoreNotFound && (code === 404 || code === 400)) {
        return { results: [] };
      }

      throw new Error(`Erro HTTP ${code}: ${text}`);
    } catch (e) {
      let stop = false;
      let errorMessage = undefined;

      switch (true) {
        case (code === 401 || code === 403):
          stop = true;
          errorMessage = `Falha de permissão/autorização (${code}) em ${url}. Detalhe: ${e.message}`;
          break;
        case (i === maxRetries - 1):
          stop = true;
          errorMessage = `Falha após ${maxRetries} tentativas na URL ${url}. Erro: ${e.message}`;
          break;
      }

      if (stop && errorMessage !== undefined) {
        cleanUp();
        throw new Error(errorMessage);
      }

      Utilities.sleep((2 ** i) * 1000);
    }
  }
}

function getPluggyApiKey() {
  const props = PropertiesService.getScriptProperties();
  const clientId = (props.getProperty("PLUGGY_CLIENT_ID") || "").trim();
  const clientSecret = (props.getProperty("PLUGGY_CLIENT_SECRET") || "").trim();

  const url = `${PLUGGY_BASE_URL}/auth`;
  const options = {
    method: "post",
    contentType: "application/json",
    payload: JSON.stringify({ clientId, clientSecret })
  };

  const data = fetchWithRetry(url, options);
  if (!data || !data.apiKey) {
    cleanUp();
    throw new Error("A resposta de /auth não retornou uma apiKey válida.");
  }
  return String(data.apiKey).trim();
}

function getItemsList(requestOptions) {
  const rawIds = PropertiesService.getScriptProperties().getProperty("MANUAL_ITEMS_IDS");
  if (rawIds) {
    const manualItemsIds = JSON.parse(rawIds);
    if (Array.isArray(manualItemsIds) && manualItemsIds.length > 0) {
      return manualItemsIds.map(id => ({ id: String(id).trim() }));
    }
  }
  const itemsData = fetchWithRetry(`${PLUGGY_BASE_URL}/v2/items`, requestOptions);
  return itemsData.results || [];
}

function getOrCreateSheet(ss, name) {
  let sheet = ss.getSheetByName(name);
  if (!sheet) {
    sheet = ss.insertSheet(name);
  }
  return sheet;
}

// ==========================================
// 3. FUNÇÃO PRINCIPAL DE SINCRONIZAÇÃO
// ==========================================
function syncFinances() {
  setupSensitiveGlobalVariables();

  try {
    const ss = SpreadsheetApp.getActiveSpreadsheet();

    // 1. Prepara as 4 Abas
    let sheetTx = ss.getSheetByName("Transacoes");
    if (!sheetTx) {
      sheetTx = ss.getSheets()[0];
      sheetTx.setName("Transacoes");
    }
    const sheetSaldos = getOrCreateSheet(ss, "Saldos");
    const sheetInvest = getOrCreateSheet(ss, "Investimentos");
    const sheetPassivos = getOrCreateSheet(ss, "Faturas_e_Emprestimos");

    if (sheetTx.getLastRow() === 0) {
      sheetTx.appendRow([
        "ID Transação", "Data", "Conta / Cartão", "Modalidade Conta",
        "Descrição", "Valor", "Fluxo (Débito/Crédito)", "Categoria", "Status", "Data de Ingestão"
      ]);
    }

    // Carrega IDs de transações existentes (Idempotência)
    const existingIds = new Set();
    const lastRow = sheetTx.getLastRow();
    if (lastRow > 1) {
      const ids = sheetTx.getRange(2, 1, lastRow - 1, 1).getValues();
      ids.forEach(row => existingIds.add(String(row[0]).trim()));
    }

    const apiKey = getPluggyApiKey();
    const requestOptions = {
      method: "get",
      contentType: "application/json",
      headers: {
        "X-API-KEY": apiKey,
        "Accept": "application/json"
      }
    };

    const items = getItemsList(requestOptions);
    if (items.length === 0) {
      Logger.log("Nenhum item encontrado para sincronizar.");
      return;
    }

    const dateFrom = new Date();
    dateFrom.setDate(dateFrom.getDate() - DAYS_TO_SYNC);
    const fromStr = dateFrom.toISOString().split("T")[0];
    const nowIso = new Date().toISOString();

    // Arrays de Snapshot
    const newTxRows = [];
    const saldosSnapshot = [[
      "ID Conta", "Nome da Conta / Cartão", "Tipo (BANK/CREDIT)", "Subtipo",
      "Saldo Atual / Fatura (R$)", "Limite Disponível (R$)", "Limite Total (R$)",
      "Fechamento Fatura", "Vencimento Fatura", "Última Sincronização"
    ]];
    const investSnapshot = [[
      "ID Investimento", "Nome do Ativo", "Tipo", "Subtipo", "Instituição / Emissor",
      "Saldo Bruto (R$)", "Valor Líquido / Resgate (R$)", "Rentabilidade Anual (%)",
      "Taxa / Indexador", "Data de Vencimento", "Status", "Última Sincronização"
    ]];
    const passivosSnapshot = [[
      "ID Registro", "Categoria (FATURA / EMPRÉSTIMO)", "Origem / Produto",
      "Data Vencimento", "Valor Total / Saldo Devedor (R$)", "Pagamento Mínimo / Parcela (R$)",
      "Taxa de Juros / CET (%)", "Status / Detalhes", "Última Sincronização"
    ]];

    // ==========================================
    // 4. COLETA DE DADOS POR ITEM E CONTA
    // ==========================================
    for (const item of items) {
      // --- A. CONTAS, TRANSAÇÕES E FATURAS DE CARTÃO (BILLS) ---
      const accountsData = fetchWithRetry(`${PLUGGY_BASE_URL}/accounts?itemId=${item.id}`, requestOptions);
      const accounts = accountsData.results || [];

      for (const acc of accounts) {
        const accName = acc.marketingName || acc.name || "Conta sem nome";
        const creditInfo = acc.creditData || {};

        // A1. Snapshot de Saldos
        saldosSnapshot.push([
          acc.id,
          accName,
          acc.type,
          acc.subtype || "",
          acc.balance !== undefined ? acc.balance : "",
          creditInfo.availableCreditLimit !== undefined ? creditInfo.availableCreditLimit : "",
          creditInfo.creditLimit !== undefined ? creditInfo.creditLimit : "",
          creditInfo.balanceCloseDate ? creditInfo.balanceCloseDate.substring(0, 10) : "",
          creditInfo.balanceDueDate ? creditInfo.balanceDueDate.substring(0, 10) : "",
          nowIso
        ]);

        // A2. Se for Cartão de Crédito, busca as Faturas (Bills)
        if (acc.type === "CREDIT") {
          const billsData = fetchWithRetry(`${PLUGGY_BASE_URL}/bills?accountId=${acc.id}`, requestOptions, 2, true);
          const bills = billsData.results || [];
          // Pega as 3 faturas mais recentes para não poluir com histórico antigo
          for (const bill of bills.slice(0, 3)) {
            passivosSnapshot.push([
              bill.id,
              "FATURA CARTÃO",
              accName,
              bill.dueDate ? bill.dueDate.substring(0, 10) : "",
              bill.totalAmount !== undefined ? bill.totalAmount : "",
              bill.minimumPaymentAmount !== undefined ? bill.minimumPaymentAmount : "",
              "",
              bill.state || "OPEN",
              nowIso
            ]);
          }
        }

        // A3. Busca Transações paginadas via cursor (/v2/transactions)
        let txUrl = `${PLUGGY_BASE_URL}/v2/transactions?accountId=${acc.id}&dateFrom=${fromStr}`;
        while (txUrl) {
          const txData = fetchWithRetry(txUrl, requestOptions);
          const transactions = txData.results || [];

          for (const tx of transactions) {
            const txId = String(tx.id).trim();
            if (!existingIds.has(txId)) {
              const isDebit = tx.type === "DEBIT" || tx.amount < 0;
              const flowLabel = isDebit ? "Saída / Gasto" : "Entrada / Receita";
              const normalizedAmount = isDebit ? -Math.abs(tx.amount) : Math.abs(tx.amount);

              newTxRows.push([
                txId,
                tx.date ? tx.date.substring(0, 10) : "",
                accName,
                acc.type,
                tx.descriptionRaw || tx.description || "",
                normalizedAmount,
                flowLabel,
                tx.category || "Sem categoria",
                tx.status || "POSTED",
                nowIso
              ]);
              existingIds.add(txId);
            }
          }
          txUrl = txData.next ? `${PLUGGY_BASE_URL}/v2/transactions${txData.next}` : null;
        }
      }

      // --- B. INVESTIMENTOS (INVESTMENTS) ---
      const investData = fetchWithRetry(`${PLUGGY_BASE_URL}/investments?itemId=${item.id}`, requestOptions, 2, true);
      const investments = investData.results || [];

      for (const inv of investments) {
        investSnapshot.push([
          inv.id,
          inv.name || "Investimento",
          inv.type || "", // FIXED_INCOME, MUTUAL_FUND, EQUITY, etc.
          inv.subtype || "",
          inv.issuer || "",
          inv.balance !== undefined ? inv.balance : "",
          inv.amountWithdrawal !== undefined ? inv.amountWithdrawal : (inv.balance || ""),
          inv.annualRate !== undefined ? inv.annualRate : "",
          inv.rateType ? `${inv.rate || ""} ${inv.rateType}` : "",
          inv.dueDate ? inv.dueDate.substring(0, 10) : "",
          inv.status || "ACTIVE",
          nowIso
        ]);
      }

      // --- C. EMPRÉSTIMOS E FINANCIAMENTOS (LOANS) ---
      const loansData = fetchWithRetry(`${PLUGGY_BASE_URL}/loans?itemId=${item.id}`, requestOptions, 2, true);
      const loans = loansData.results || [];

      for (const loan of loans) {
        const saldoDevedor = loan.payments && loan.payments.totalRemainingAmount !== undefined
          ? loan.payments.totalRemainingAmount
          : (loan.contractAmount || "");

        passivosSnapshot.push([
          loan.id,
          "EMPRÉSTIMO / FINANCIAMENTO",
          loan.productName || loan.type || "Contrato",
          loan.dueDate ? loan.dueDate.substring(0, 10) : "",
          saldoDevedor,
          "",
          loan.CET !== undefined ? loan.CET : "",
          loan.installments && loan.installments.remainingNumberOfInstallments !== undefined
            ? `${loan.installments.remainingNumberOfInstallments} parcelas restantes`
            : "Ativo",
          nowIso
        ]);
      }
    }

    // ==========================================
    // 5. ESCRITA NAS ABAS DO GOOGLE SHEETS
    // ==========================================
    if (newTxRows.length > 0) {
      sheetTx.getRange(sheetTx.getLastRow() + 1, 1, newTxRows.length, newTxRows[0].length).setValues(newTxRows);
      sheetTx.getRange(2, 1, sheetTx.getLastRow() - 1, sheetTx.getLastColumn()).sort({ column: 2, ascending: false });
      Logger.log(`Sucesso: ${newTxRows.length} novas transações adicionadas.`);
    }

    sheetSaldos.clearContents();
    sheetSaldos.getRange(1, 1, saldosSnapshot.length, saldosSnapshot[0].length).setValues(saldosSnapshot);

    sheetInvest.clearContents();
    sheetInvest.getRange(1, 1, investSnapshot.length, investSnapshot[0].length).setValues(investSnapshot);

    sheetPassivos.clearContents();
    sheetPassivos.getRange(1, 1, passivosSnapshot.length, passivosSnapshot[0].length).setValues(passivosSnapshot);

    Logger.log("Sincronização completa (Transações, Saldos, Investimentos, Faturas e Empréstimos).");

  } finally {
    cleanUp();
  }
}