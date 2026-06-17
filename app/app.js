const categoriesByKind = {
  Discretionary: [
    "Discretionary spending pool",
    "Restaurants / coffee",
    "Fun / summer plans",
    "Shopping / personal",
    "Miscellaneous"
  ],
  Essentials: ["Rent", "Groceries", "Gas", "Storage unit", "Health", "Personal"],
  Savings: ["Savings transfer", "Emergency fund", "Long-term savings"],
  Emergency: ["Car", "Medical", "Family", "Unexpected bill", "Other emergency"],
};

const state = {
  expenses: JSON.parse(localStorage.getItem("budgetTap.expenses") || "[]"),
  settings: JSON.parse(
    localStorage.getItem("budgetTap.settings") ||
      '{"endpoint":"","sheetName":"Expense Log"}'
  ),
};

const form = document.querySelector("#expenseForm");
const category = document.querySelector("#category");
const recentList = document.querySelector("#recentList");
const statusEl = document.querySelector("#status");
const settingsDialog = document.querySelector("#settingsDialog");
const endpointInput = document.querySelector("#endpoint");
const sheetNameInput = document.querySelector("#sheetName");

function money(value) {
  return Number(value).toLocaleString(undefined, {
    style: "currency",
    currency: "USD",
  });
}

function selected(name) {
  return document.querySelector(`input[name="${name}"]:checked`)?.value || "";
}

function createId() {
  if (crypto?.randomUUID) return crypto.randomUUID();
  return `expense-${Date.now()}-${Math.random().toString(36).slice(2, 10)}`;
}

function setStatus(message, tone = "") {
  statusEl.textContent = message;
  statusEl.className = `status ${tone}`.trim();
}

function updateCategories(kind = selected("kind")) {
  category.innerHTML = "";
  for (const item of categoriesByKind[kind] || []) {
    const option = document.createElement("option");
    option.value = item;
    option.textContent = item;
    category.append(option);
  }
}

function saveLocal() {
  localStorage.setItem("budgetTap.expenses", JSON.stringify(state.expenses));
  localStorage.setItem("budgetTap.settings", JSON.stringify(state.settings));
}

function renderRecent() {
  recentList.innerHTML = "";
  const recent = state.expenses.slice(0, 8);

  if (recent.length === 0) {
    const empty = document.createElement("li");
    empty.innerHTML = `<span class="recent-meta">No expenses yet.</span>`;
    recentList.append(empty);
    return;
  }

  for (const expense of recent) {
    const item = document.createElement("li");
    item.innerHTML = `
      <span class="recent-title">${expense.merchant || expense.category}</span>
      <span class="recent-amount">${money(expense.amount)}</span>
      <span class="recent-meta">${expense.kind} · ${expense.card} · ${expense.category}</span>
      <span class="recent-meta">${expense.synced ? "Synced" : "Queued"}</span>
    `;
    recentList.append(item);
  }
}
async function postExpense(expense) {
  if (!state.settings.endpoint) return false;

  try {
    const res = await fetch(state.settings.endpoint, {
      method: "POST",
      body: JSON.stringify({
        sheetName: state.settings.sheetName,
        expense,
      }),
    });

    const text = await res.text();
    console.log("SYNC RESPONSE:", text);

    return res.ok;
  } catch (err) {
    console.log("SYNC ERROR:", err);
    return false;
  }
}

async function syncQueued() {
  const queued = state.expenses.filter((e) => !e.synced);

  if (queued.length === 0) {
    setStatus("Nothing queued.", "good");
    return;
  }

  if (!state.settings.endpoint) {
    setStatus("Add your Google Apps Script URL in settings first.", "warn");
    settingsDialog.showModal();
    return;
  }

  let synced = 0;

  for (const expense of queued) {
    const ok = await postExpense(expense);
    if (ok) {
      expense.synced = true;
      synced++;
    }
  }

  saveLocal();
  renderRecent();

  setStatus(
    synced
      ? `Synced ${synced} expense${synced === 1 ? "" : "s"}.`
      : "Sync failed. Still queued.",
    synced ? "good" : "bad"
  );
}

form.addEventListener("change", (event) => {
  if (event.target.name === "kind") updateCategories(event.target.value);
});

form.addEventListener("submit", async (event) => {
  event.preventDefault();

  const data = new FormData(form);
  const amount = Number(data.get("amount"));

  if (!Number.isFinite(amount) || amount <= 0) {
    setStatus("Enter an amount greater than zero.", "bad");
    return;
  }

  const expense = {
    id: createId(),
    createdAt: new Date().toISOString(),
    date: new Date().toLocaleDateString("en-CA"),
    amount,
    merchant: String(data.get("merchant") || "").trim(),
    kind: selected("kind"),
    category: data.get("category"),
    card: selected("card"),
    note: String(data.get("note") || "").trim(),
    synced: false,
  };

  state.expenses.unshift(expense);
  saveLocal();
  renderRecent();

  form.reset();
  document.querySelector('input[name="kind"][value="Discretionary"]').checked = true;
  document.querySelector('input[name="card"][value="Cap 1"]').checked = true;
  updateCategories("Discretionary");
  document.querySelector("#amount").focus();

  const ok = await postExpense(expense);
  expense.synced = ok;

  saveLocal();
  renderRecent();

  setStatus(
    ok ? "Sent to Google Sheets." : "Saved locally (queued for sync).",
    ok ? "good" : "warn"
  );
});

document.querySelector("#settingsButton").addEventListener("click", () => {
  endpointInput.value = state.settings.endpoint;
  sheetNameInput.value = state.settings.sheetName;
  settingsDialog.showModal();
});

document.querySelector("#saveSettings").addEventListener("click", () => {
  state.settings.endpoint = endpointInput.value.trim();
  state.settings.sheetName =
    sheetNameInput.value.trim() || "Expense Log";

  saveLocal();
  settingsDialog.close();

  setStatus("Settings saved.", "good");
});

document.querySelector("#syncButton").addEventListener("click", syncQueued);

updateCategories();
renderRecent();
