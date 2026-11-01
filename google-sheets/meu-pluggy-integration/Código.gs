// ==========================================
// 1. CONFIGURAÇÕES GLOBAIS E CREDENCIAIS
// ==========================================
// Referência da API: https://docs.pluggy.ai/en/reference
const PLUGGY_BASE_URL = "https://api.pluggy.ai";

// Quantos dias para trás buscar a cada sincronização
const DAYS_TO_SYNC = 90;

// Quantas faturas mais recentes exibir por cartão
const BILLS_TO_SHOW = 3;

// Tolerância (R$) para arredondamentos ao comparar valores pagos x devidos
const AMOUNT_TOLERANCE = 0.05;

const CARD_PAYMENT_LABEL = "Pagamento de Fatura / Abatimento";

function setupCredentials() {
  PropertiesService.getScriptProperties().setProperties({
    "PLUGGY_CLIENT_ID": "SEU_CLIENT_ID_AQUI",
    "PLUGGY_CLIENT_SECRET": "SEU_CLIENT_SECRET_AQUI"
  });
}

function setupManualItemsIds() {
  // Coloque os IDs únicos das suas conexões (sem repetir IDs)
  const ids = [
    "550e8400-e29b-41d4-a716-446655440000", // Exemplos de UUIDs
    "6ba7b810-9dad-41d1-80b4-00c04fd430c8"
    // "ID_ITEM_3",
    // "ID_ITEM_4",
    // "ID_ITEM_5" (Máximo de Conexões permitidas no plano gratuito)
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
// 2. CLIENTE HTTP COM RETRY & HELPERS
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

/**
 * Nos endpoints de cursor (/v2/transactions, /v2/items), `next` é uma query string
 * pronta para ser anexada ao path (ex: GET /v2/transactions{next}), ou null no fim.
 */
function buildNextCursorUrl(endpointPath, next) {
  return next ? `${PLUGGY_BASE_URL}${endpointPath}${next}` : null;
}

function toDateStr(value) {
  return value ? String(value).substring(0, 10) : "";
}

function sumAmounts(list) {
  if (!Array.isArray(list)) return 0;
  return list.reduce((acc, entry) => acc + Number(entry.amount || 0), 0);
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
      const uniqueIds = [...new Set(manualItemsIds.map(id => String(id).trim()).filter(Boolean))];
      return uniqueIds.map(id => ({ id }));
    }
  }

  // Fallback paginado em /v2/items (items-list-by-cursor)
  const allItems = [];
  let itemsUrl = `${PLUGGY_BASE_URL}/v2/items`;
  while (itemsUrl) {
    const itemsData = fetchWithRetry(itemsUrl, requestOptions);
    allItems.push(...(itemsData.results || []));
    itemsUrl = buildNextCursorUrl("/v2/items", itemsData.next);
  }
  return allItems;
}

function getOrCreateSheet(ss, name) {
  let sheet = ss.getSheetByName(name);
  if (!sheet) {
    sheet = ss.insertSheet(name);
  }
  return sheet;
}

// ==========================================
// 3. REGRAS DE NEGÓCIO (TRANSAÇÕES, FATURAS E EMPRÉSTIMOS)
// ==========================================

/**
 * Pagamento de fatura em conta CREDIT.
 * Open Finance: `operationType === "PAGAMENTO_FATURA"` (inclusive o "Pagamento recebido"
 * que a própria Pluggy sintetiza a partir da fatura). Conectores diretos não trazem
 * `operationType`, então caímos para categoria/descrição.
 * Estornos e cashback também são `type: CREDIT`, mas NÃO são pagamento de fatura.
 */
function isCardBillPayment(tx) {
  if (tx.operationType) return tx.operationType === "PAGAMENTO_FATURA";
  if (tx.type !== "CREDIT") return false;

  const desc = (tx.descriptionRaw || tx.description || "").toLowerCase();
  const cat = (tx.category || "").toLowerCase();
  return cat === "credit card payment" ||
    desc.includes("pagamento recebido") ||
    desc.includes("pagamento de fatura") ||
    desc.includes("pagamento fatura");
}

/**
 * A Pluggy normaliza `type` do ponto de vista do titular, inclusive em cartões:
 * compras são DEBIT e pagamentos/estornos são CREDIT.
 */
function classifyTransaction(tx, accountType) {
  const absAmount = Math.abs(Number(tx.amount || 0));

  if (accountType === "CREDIT" && isCardBillPayment(tx)) {
    return { flowLabel: CARD_PAYMENT_LABEL, normalizedAmount: absAmount };
  }

  const isDebit = tx.type ? tx.type === "DEBIT" : Number(tx.amount || 0) < 0;
  return {
    flowLabel: isDebit ? "Saída / Gasto" : "Entrada / Receita",
    normalizedAmount: isDebit ? -absAmount : absAmount
  };
}

/**
 * Pagamentos de fatura feitos dentro da janela da fatura N:
 * - com `billClosingDate`: [fechamento N, fechamento N+1)
 * - sem `billClosingDate` (comum em Open Finance): (vencimento N-1, vencimento N]
 */
function sumCardPaymentsInBillWindow(bill, newerBill, olderBill, cardPayments) {
  const closingDate = toDateStr(bill.billClosingDate);
  const previousDueDate = toDateStr(olderBill && olderBill.dueDate);
  if (!closingDate && !previousDueDate) return 0;

  const nextClosingDate = toDateStr(newerBill && newerBill.billClosingDate);
  const dueDate = toDateStr(bill.dueDate);

  return cardPayments
    .filter(payment => {
      if (closingDate) {
        return payment.date >= closingDate && (!nextClosingDate || payment.date < nextClosingDate);
      }
      return payment.date > previousDueDate && (!dueDate || payment.date <= dueDate);
    })
    .reduce((acc, payment) => acc + payment.amount, 0);
}

/**
 * O schema de Bill NÃO tem campo de status. O valor pago é o maior entre 3 evidências:
 * 1. `payments[]` da própria fatura: na prática (Nubank, C6, Mercado Pago via Open Finance)
 *    a instituição registra o pagamento na fatura que foi paga.
 * 2. Regra do guia oficial (https://docs.pluggy.ai/en/docs/products/credit-card-bills):
 *    totalAmount(N) + sum(financeCharges(N+1)) = sum(payments(N+1))
 * 3. Transações de pagamento do cartão dentro da janela da fatura (cobre a fatura
 *    mais recente quando o pagamento ainda não entrou em `payments[]`).
 */
function getAmountPaidTowardsBill(bill, newerBill, olderBill, cardPayments) {
  const ownPayments = sumAmounts(bill.payments);
  const nextCyclePayments = newerBill
    ? sumAmounts(newerBill.payments) - sumAmounts(newerBill.financeCharges)
    : 0;
  const transactionPayments = sumCardPaymentsInBillWindow(bill, newerBill, olderBill, cardPayments);

  return Math.max(ownPayments, nextCyclePayments, transactionPayments);
}

function evaluateBillStatus(bill, newerBill, olderBill, cardPayments, todayStr) {
  const billTotal = Number(bill.totalAmount || 0);
  if (billTotal <= 0) return "PAID (Zerada)";

  const paidAmount = getAmountPaidTowardsBill(bill, newerBill, olderBill, cardPayments);
  if (paidAmount >= billTotal - AMOUNT_TOLERANCE) {
    return `PAID (Pago R$ ${paidAmount.toFixed(2)})`;
  }

  const dueDate = toDateStr(bill.dueDate);
  const isOverdue = Boolean(dueDate) && dueDate < todayStr;

  if (paidAmount > AMOUNT_TOLERANCE) {
    const remaining = (billTotal - paidAmount).toFixed(2);
    return `${isOverdue ? "OVERDUE" : "PARTIAL"} (Pago R$ ${paidAmount.toFixed(2)} | Resta R$ ${remaining})`;
  }

  return isOverdue ? "OVERDUE" : "OPEN";
}

const LOAN_TERM_UNITS = {
  DAY: "dia(s)",
  WEEK: "semana(s)",
  MONTH: "mês(es)",
  YEAR: "ano(s)"
};

/**
 * Campos reais de Loan (padrão Open Finance):
 * - installments: { contractRemainingNumber, typeContractRemaining, pastDueInstallments, ... }
 * - payments: { contractOutstandingBalance, releases[] }
 */
function describeLoanStatus(installments, outstandingBalance) {
  const pastDue = Number(installments.pastDueInstallments || 0);
  if (pastDue > 0) return `OVERDUE (${pastDue} parcela(s) em atraso)`;

  if (outstandingBalance === 0) return "PAID";

  const remaining = installments.contractRemainingNumber;
  const unit = LOAN_TERM_UNITS[installments.typeContractRemaining];
  if (remaining !== undefined && remaining !== null && unit) {
    return `OPEN (Prazo restante: ${remaining} ${unit})`;
  }

  return "OPEN";
}

function getLastRegularLoanPayment(loanPayments) {
  const releases = Array.isArray(loanPayments.releases) ? loanPayments.releases : [];
  return releases
    .filter(release => !release.isOverParcelPayment)
    .sort((a, b) => String(b.paidDate || "").localeCompare(String(a.paidDate || "")))[0];
}

// ==========================================
// 4. FUNÇÃO PRINCIPAL DE SINCRONIZAÇÃO
// ==========================================
function syncFinances() {
  setupSensitiveGlobalVariables();

  try {
    const ss = SpreadsheetApp.getActiveSpreadsheet();

    const sheetTx = getOrCreateSheet(ss, "Transacoes");
    const sheetSaldos = getOrCreateSheet(ss, "Saldos");
    const sheetInvest = getOrCreateSheet(ss, "Investimentos");
    const sheetPassivos = getOrCreateSheet(ss, "Faturas_e_Emprestimos");

    const txHeaders = [
      "ID Transação", "Data", "Conta / Cartão", "Modalidade Conta",
      "Descrição", "Valor", "Fluxo (Débito/Crédito)", "Categoria", "Status", "Data de Ingestão"
    ];

    if (sheetTx.getLastRow() === 0) {
      sheetTx.appendRow(txHeaders);
    }

    // Mapeia transações existentes para atualizar PENDING -> POSTED e corrigir sinais antigos
    const existingTxMap = new Map();
    const lastRow = sheetTx.getLastRow();
    let existingSheetValues = [];
    if (lastRow > 1) {
      existingSheetValues = sheetTx.getRange(2, 1, lastRow - 1, 10).getValues();
      existingSheetValues.forEach((row, idx) => {
        const id = String(row[0]).trim();
        if (id) {
          existingTxMap.set(id, { index: idx, data: row });
        }
      });
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
    const todayStr = Utilities.formatDate(new Date(), Session.getScriptTimeZone(), "yyyy-MM-dd");

    // Sets de deduplicação por entidade
    const seenAccountIds = new Set();
    const seenBillIds = new Set();
    const seenInvestIds = new Set();
    const seenLoanIds = new Set();

    const newTxRows = [];
    let updatedExistingCount = 0;

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
    // 5. COLETA DE DADOS POR ITEM E CONTA
    // ==========================================
    for (const item of items) {
      // --- A. CONTAS (GET /accounts) ---
      const accountsData = fetchWithRetry(`${PLUGGY_BASE_URL}/accounts?itemId=${item.id}`, requestOptions);
      const accounts = accountsData.results || [];

      for (const acc of accounts) {
        if (seenAccountIds.has(acc.id)) continue;
        seenAccountIds.add(acc.id);

        const accName = acc.marketingName || acc.name || "Conta sem nome";
        const creditInfo = acc.creditData || {};

        saldosSnapshot.push([
          acc.id,
          accName,
          acc.type,
          acc.subtype || "",
          acc.balance ?? "",
          creditInfo.availableCreditLimit ?? "",
          creditInfo.creditLimit ?? "",
          toDateStr(creditInfo.balanceCloseDate),
          toDateStr(creditInfo.balanceDueDate),
          nowIso
        ]);

        const cardPayments = [];
        const cardPaymentKeys = new Set();

        // --- A2. TRANSAÇÕES (GET /v2/transactions) ---
        // Na v2 (cursor) o filtro de data é `dateFrom`. Não usar GET /transactions (v1, deprecated, que usa `from`/`page`)
        let txUrl = `${PLUGGY_BASE_URL}/v2/transactions?accountId=${acc.id}&dateFrom=${fromStr}`;
        while (txUrl) {
          const txData = fetchWithRetry(txUrl, requestOptions);
          const transactions = txData.results || [];

          for (const tx of transactions) {
            const txId = String(tx.id).trim();
            const txDate = toDateStr(tx.date);
            const txDesc = tx.descriptionRaw || tx.description || "";
            const txCategory = tx.category || "Sem categoria";
            const txStatus = tx.status || "POSTED";

            const { flowLabel, normalizedAmount } = classifyTransaction(tx, acc.type);

            // O mesmo pagamento pode vir 2x: o da instituição (ex: "Pagamento Fatura QR CODE", PENDING)
            // e o "Pagamento recebido" que a Pluggy sintetiza a partir da fatura
            const paymentKey = `${txDate}|${normalizedAmount.toFixed(2)}`;
            if (flowLabel === CARD_PAYMENT_LABEL && !cardPaymentKeys.has(paymentKey)) {
              cardPaymentKeys.add(paymentKey);
              cardPayments.push({ date: txDate, amount: normalizedAmount });
            }

            const rowPayload = [
              txId,
              txDate,
              accName,
              acc.type,
              txDesc,
              normalizedAmount,
              flowLabel,
              txCategory,
              txStatus,
              nowIso
            ];

            if (!existingTxMap.has(txId)) {
              newTxRows.push(rowPayload);
              existingTxMap.set(txId, { index: -1, data: rowPayload });
              continue;
            }

            const existingEntry = existingTxMap.get(txId);
            if (existingEntry.index === -1) continue;

            const oldRow = existingEntry.data;
            const hasChanged =
              String(oldRow[8]) !== txStatus ||
              String(oldRow[6]) !== flowLabel ||
              Math.abs(Number(oldRow[5]) - normalizedAmount) > 0.001;

            if (hasChanged) {
              existingSheetValues[existingEntry.index] = rowPayload;
              updatedExistingCount++;
            }
          }
          txUrl = buildNextCursorUrl("/v2/transactions", txData.next);
        }

        // --- A3. FATURAS DE CARTÃO (GET /bills) ---
        if (acc.type === "CREDIT") {
          const billsData = fetchWithRetry(`${PLUGGY_BASE_URL}/bills?accountId=${acc.id}`, requestOptions, 2, true);
          // A API não garante ordenação: ordena da mais recente para a mais antiga
          const bills = (billsData.results || [])
            .slice()
            .sort((a, b) => toDateStr(b.dueDate).localeCompare(toDateStr(a.dueDate)));

          bills.slice(0, BILLS_TO_SHOW).forEach((bill, i) => {
            if (seenBillIds.has(bill.id)) return;
            seenBillIds.add(bill.id);

            const newerBill = bills[i - 1] || null;
            const olderBill = bills[i + 1] || null;
            const computedStatus = evaluateBillStatus(bill, newerBill, olderBill, cardPayments, todayStr);
            const financeChargesSum = sumAmounts(bill.financeCharges);

            passivosSnapshot.push([
              bill.id,
              "FATURA CARTÃO",
              accName,
              toDateStr(bill.dueDate),
              bill.totalAmount ?? "",
              bill.minimumPaymentAmount ?? "",
              financeChargesSum > 0 ? `Encargos: R$ ${financeChargesSum.toFixed(2)}` : "",
              computedStatus,
              nowIso
            ]);
          });
        }
      }

      // --- B. INVESTIMENTOS (GET /investments) ---
      const investData = fetchWithRetry(`${PLUGGY_BASE_URL}/investments?itemId=${item.id}`, requestOptions, 2, true);
      const investments = investData.results || [];

      for (const inv of investments) {
        if (seenInvestIds.has(inv.id)) continue;
        seenInvestIds.add(inv.id);

        const issuerName = inv.issuer || (inv.institution && inv.institution.name) || "";
        const annualYield = inv.annualRate ?? inv.lastTwelveMonthsRate ?? "";

        // Formata a taxa acordada (ex: "100% CDI" ou "IPCA + 6%")
        const rateParts = [];
        if (inv.rate !== undefined && inv.rate !== null) rateParts.push(`${inv.rate}%`);
        if (inv.rateType) rateParts.push(inv.rateType);
        if (Number(inv.fixedAnnualRate) > 0) rateParts.push(`+ ${inv.fixedAnnualRate}% a.a.`);

        investSnapshot.push([
          inv.id,
          inv.name || "Investimento",
          inv.type || "",
          inv.subtype || "",
          issuerName,
          inv.amount ?? inv.balance ?? "",
          inv.amountWithdrawal ?? inv.balance ?? "",
          annualYield,
          rateParts.join(" "),
          toDateStr(inv.dueDate),
          inv.status || "ACTIVE",
          nowIso
        ]);
      }

      // --- C. EMPRÉSTIMOS E FINANCIAMENTOS (GET /loans) ---
      const loansData = fetchWithRetry(`${PLUGGY_BASE_URL}/loans?itemId=${item.id}`, requestOptions, 2, true);
      const loans = loansData.results || [];

      for (const loan of loans) {
        if (seenLoanIds.has(loan.id)) continue;
        seenLoanIds.add(loan.id);

        const loanPayments = loan.payments || {};
        const loanInstallments = loan.installments || {};
        const outstandingBalance = loanPayments.contractOutstandingBalance ?? null;
        const lastPayment = getLastRegularLoanPayment(loanPayments);

        passivosSnapshot.push([
          loan.id,
          "EMPRÉSTIMO / FINANCIAMENTO",
          loan.productName || loan.type || "Contrato",
          toDateStr(loan.dueDate),
          outstandingBalance ?? "",
          lastPayment ? lastPayment.paidAmount ?? "" : "",
          loan.CET ?? "",
          describeLoanStatus(loanInstallments, outstandingBalance),
          nowIso
        ]);
      }
    }

    // ==========================================
    // 6. ESCRITA E ATUALIZAÇÃO NO GOOGLE SHEETS
    // ==========================================
    if (updatedExistingCount > 0 && existingSheetValues.length > 0) {
      sheetTx.getRange(2, 1, existingSheetValues.length, 10).setValues(existingSheetValues);
      Logger.log(`Atualizadas ${updatedExistingCount} transações existentes.`);
    }

    if (newTxRows.length > 0) {
      sheetTx.getRange(sheetTx.getLastRow() + 1, 1, newTxRows.length, newTxRows[0].length).setValues(newTxRows);
      Logger.log(`Sucesso: ${newTxRows.length} novas transações adicionadas.`);
    }

    if (sheetTx.getLastRow() > 1) {
      sheetTx.getRange(2, 1, sheetTx.getLastRow() - 1, sheetTx.getLastColumn()).sort({ column: 2, ascending: false });
    }

    sheetSaldos.clearContents();
    sheetSaldos.getRange(1, 1, saldosSnapshot.length, saldosSnapshot[0].length).setValues(saldosSnapshot);

    sheetInvest.clearContents();
    sheetInvest.getRange(1, 1, investSnapshot.length, investSnapshot[0].length).setValues(investSnapshot);

    sheetPassivos.clearContents();
    sheetPassivos.getRange(1, 1, passivosSnapshot.length, passivosSnapshot[0].length).setValues(passivosSnapshot);

    Logger.log("Sincronização concluída.");

  } finally {
    cleanUp();
  }
}
