function sleep(ms) {
  return new Promise(resolve => setTimeout(resolve, ms));
}

function normalizeText(text) {
  return (text || "")
    .replace(/\u00a0/g, " ")
    .replace(/\s+/g, " ")
    .trim()
    .toLowerCase();
}

function getCurrentDescription() {
  const now = new Date();
  const month = now.toLocaleString("en-US", { month: "long" });
  const year = now.getFullYear();
  return `Bloomberg Access - ${month} ${year}`;
}

function logStep(message) {
  console.log("[Bloomberg Filler]", message);
}

function findRowByLabel(labelText) {
  const target = normalizeText(labelText);
  const labels = Array.from(document.querySelectorAll("label"));
  for (const label of labels) {
    if (normalizeText(label.textContent) === target) {
      return label.closest("tr");
    }
  }
  return null;
}

function findButtonByText(text) {
  const target = normalizeText(text);
  return Array.from(document.querySelectorAll("button")).find(
    btn => normalizeText(btn.textContent) === target
  );
}

function setInputValue(input, value) {
  if (!input) return;

  input.focus();

  const proto =
    input.tagName === "TEXTAREA"
      ? HTMLTextAreaElement.prototype
      : HTMLInputElement.prototype;

  const desc = Object.getOwnPropertyDescriptor(proto, "value");
  if (desc?.set) {
    desc.set.call(input, value);
  } else {
    input.value = value;
  }

  input.dispatchEvent(new Event("input", { bubbles: true }));
  input.dispatchEvent(new Event("change", { bubbles: true }));
}

async function waitFor(getter, timeoutMs = 12000, intervalMs = 250) {
  const start = Date.now();
  while (Date.now() - start < timeoutMs) {
    const value = getter();
    if (value) return value;
    await sleep(intervalMs);
  }
  return null;
}

async function setPlainField(label, value) {
  logStep(`Setting plain field: ${label} -> ${value}`);
  const row = findRowByLabel(label);
  if (!row) throw new Error(`Row not found for ${label}`);

  const input = row.querySelector("textarea, input[type='text'], input:not([type])");
  if (!input) throw new Error(`Input not found for ${label}`);

  setInputValue(input, value);
  input.dispatchEvent(new Event("blur", { bubbles: true }));
  await sleep(300);
}

async function selectChooserExact(label, searchText, exactOptionText) {
  logStep(`Selecting chooser: ${label} -> ${exactOptionText}`);
  const row = findRowByLabel(label);
  if (!row) throw new Error(`Row not found for ${label}`);

  const input = row.querySelector("input.w-chInput");
  if (!input) throw new Error(`Chooser input not found for ${label}`);

  const wrapper = row.querySelector(".w-chWrapper");
  const menuId = wrapper?.querySelector(".w-chWrap")?.getAttribute("awmenuid");
  if (!menuId) throw new Error(`Chooser menu id not found for ${label}`);

  const exact = normalizeText(exactOptionText);

  // Clear then type
  setInputValue(input, "");
  await sleep(150);
  setInputValue(input, searchText);
  await sleep(400);

  // Open dropdown explicitly
  const arrow =
    row.querySelector(".w-chPullDownDiv") ||
    row.querySelector(".w-chWrapRight a") ||
    row.querySelector(".w-chWrapRight");

  if (arrow) {
    arrow.click();
  } else {
    input.click();
  }

  const menu = await waitFor(() => document.getElementById(menuId), 5000, 200);
  if (!menu) throw new Error(`Chooser menu not found for ${label}`);

  // Wait for exact option
  const option = await waitFor(() => {
    const options = Array.from(menu.querySelectorAll("[role='option'], .w-pmi-item"));
    return options.find(opt => normalizeText(opt.textContent) === exact);
  }, 10000, 200);

  if (!option) {
    throw new Error(`Exact option "${exactOptionText}" not found for ${label}`);
  }

  option.scrollIntoView({ block: "nearest" });
  await sleep(150);
  option.click();
  await sleep(1200);

  input.dispatchEvent(new Event("change", { bubbles: true }));
  input.dispatchEvent(new Event("blur", { bubbles: true }));
  await sleep(1000);
}

function isElementVisible(el) {
  if (!el) return false;
  const style = window.getComputedStyle(el);
  if (style.display === "none" || style.visibility === "hidden" || style.opacity === "0") {
    return false;
  }
  const rect = el.getBoundingClientRect();
  return rect.width > 0 && rect.height > 0;
}

async function selectContactViaSelectButtons() {
  const contactRow = findRowByLabel("Contact:");
  if (!contactRow) {
    throw new Error("Row not found for Contact:");
  }

  const rowSelectButton = Array.from(contactRow.querySelectorAll("button, a, [role='button']")).find(
    el => normalizeText(el.textContent) === "select"
  );

  if (!rowSelectButton) {
    throw new Error('No "Select" button found in Contact row');
  }

  // Some flows require opening Contact chooser and then confirming with Select.
  for (let i = 0; i < 2; i++) {
    rowSelectButton.click();
    await sleep(700);

    const confirmSelect = await waitFor(() => {
      const candidates = Array.from(document.querySelectorAll("button, a, [role='button']"));
      return candidates.find(el => {
        const text = normalizeText(el.textContent);
        const isInContactRow = contactRow.contains(el);
        return text === "select" && !isInContactRow && isElementVisible(el);
      });
    }, 4000, 200);

    if (!confirmSelect) {
      // If no secondary Select appears, first click may already complete the selection.
      break;
    }

    // If there is a result grid/list, pick the first selectable row before confirming.
    const firstResult = Array.from(
      document.querySelectorAll("[role='row'], .w-tbl-row, .w-pmi-item, li")
    ).find(el => {
      if (!isElementVisible(el)) return false;
      const text = normalizeText(el.textContent);
      return text && !text.includes("search more") && text !== "select";
    });

    if (firstResult) {
      firstResult.click();
      await sleep(200);
    }

    confirmSelect.click();
    await sleep(900);
  }
}

async function selectContactForVendor() {
  logStep("Selecting Contact for Vendor");

  const expectedContact = "0002190112 (2190112 BLOOMBERG FINANCE L P)";

  try {
    await selectChooserExact("Contact:", "2190112", expectedContact);
    return;
  } catch (chooserErr) {
    logStep(`Contact chooser selection fallback: ${chooserErr.message}`);
  }

  await selectContactViaSelectButtons();
}

async function openAccountTypeDropdown() {
  const row = findRowByLabel("Account Type:");
  if (!row) throw new Error("Row not found for Account Type:");

  const combo = row.querySelector("[role='combobox'].w-dropdown");
  if (!combo) throw new Error("Dropdown not found for Account Type:");

  combo.click();
  await sleep(400);

  const comboId = combo.getAttribute("id");
  const itemsContainer = comboId ? document.getElementById(`Items_${comboId}`) : null;
  if (!itemsContainer) throw new Error("Dropdown items not found for Account Type:");

  return { combo, itemsContainer };
}

async function selectAccountTypeWbs() {
  logStep("Selecting Account Type -> WBS element");

  const ready = await waitFor(async () => {
    const row = findRowByLabel("Account Type:");
    if (!row) return null;

    const combo = row.querySelector("[role='combobox'].w-dropdown");
    if (!combo) return null;

    combo.click();
    await sleep(250);

    const comboId = combo.getAttribute("id");
    const itemsContainer = comboId ? document.getElementById(`Items_${comboId}`) : null;
    if (!itemsContainer) return null;

    const options = Array.from(
      itemsContainer.querySelectorAll("[role='option'], .w-dropdown-item")
    );

    return options.length > 1 ? { combo, itemsContainer } : null;
  }, 20000, 500);

  if (!ready) {
    throw new Error("Account Type options did not populate");
  }

  const { combo, itemsContainer } = ready;
  const target = normalizeText("WBS element");

  for (let i = 0; i < 40; i++) {
    const options = Array.from(
      itemsContainer.querySelectorAll("[role='option'], .w-dropdown-item")
    );

    let match = options.find(opt => normalizeText(opt.textContent) === target);
    if (!match) {
      match = options.find(opt => normalizeText(opt.textContent).includes(target));
    }

    if (match) {
      match.scrollIntoView({ block: "nearest" });
      await sleep(150);
      match.click();
      await sleep(1000);
      return;
    }

    itemsContainer.scrollTop += 120;
    await sleep(250);
  }

  combo.click();
  throw new Error('Option "WBS element" not found in Account Type dropdown');
}

async function clickButton(text) {
  logStep(`Clicking button: ${text}`);
  const button = findButtonByText(text);
  if (!button) throw new Error(`Button not found: ${text}`);
  button.click();
  await sleep(1200);
}

function getProceedToCheckoutElement() {
  return (
    document.getElementById("_xxfxrb") ||
    Array.from(document.querySelectorAll("a, button")).find(el =>
      normalizeText(el.textContent).includes("proceed to checkout")
    )
  );
}

async function proceedAfterAddToCart() {
  logStep("Waiting for Proceed to Checkout");
  const proceed = await waitFor(getProceedToCheckoutElement, 15000, 300);
  if (!proceed) {
    throw new Error("Proceed to Checkout not found after Add to Cart");
  }
  proceed.click();
  await sleep(1500);
}

async function fillBloombergForm() {
  const results = [];

  try {
    await setPlainField("Full Description:", getCurrentDescription());
    results.push("Full Description");
  } catch (e) {
    results.push(`Full Description failed: ${e.message}`);
  }

  // Vendor first
  try {
    await selectChooserExact(
      "Vendor:",
      "bloomberg finance",
      "0002190112 (2190112 BLOOMBERG FINANCE L P)"
    );
    results.push("Vendor");
  } catch (e) {
    results.push(`Vendor failed: ${e.message}`);
    return results;
  }

  try {
    await selectContactForVendor();
    results.push("Contact");
  } catch (e) {
    results.push(`Contact failed: ${e.message}`);
    return results;
  }

  // Commodity next
  try {
    await selectChooserExact(
      "Commodity Code:",
      "D03",
      "D03-NETWORK SERVICES"
    );
    results.push("Commodity Code");
  } catch (e) {
    results.push(`Commodity Code failed: ${e.message}`);
    return results;
  }

  try {
    await selectAccountTypeWbs();
    results.push("Account Type");
  } catch (e) {
    results.push(`Account Type failed: ${e.message}`);
  }

  try {
    await setPlainField("Quantity:", "1");
    results.push("Quantity");
  } catch (e) {
    results.push(`Quantity failed: ${e.message}`);
  }

  try {
    await selectChooserExact("Unit of Measure:", "each", "each");
    results.push("Unit of Measure");
  } catch (e) {
    results.push(`Unit of Measure failed: ${e.message}`);
  }

  try {
    await setPlainField("Price:", "2035.98");
    results.push("Price");
  } catch (e) {
    results.push(`Price failed: ${e.message}`);
  }

  try {
    await clickButton("Update Amount");
    results.push("Update Amount");
  } catch (e) {
    results.push(`Update Amount failed: ${e.message}`);
  }

  try {
    await clickButton("Add to Cart");
    results.push("Add to Cart");
  } catch (e) {
    results.push(`Add to Cart failed: ${e.message}`);
    return results;
  }

  try {
    await proceedAfterAddToCart();
    results.push("Proceed to Checkout");
  } catch (e) {
    results.push(`Proceed to Checkout failed: ${e.message}`);
  }

  return results;
}

chrome.runtime.onMessage.addListener((message, sender, sendResponse) => {
  if (message?.type !== "FILL_BLOOMBERG_FORM") return;

  (async () => {
    try {
      const results = await fillBloombergForm();
      const failures = results.filter(r => r.includes("failed:"));

      if (failures.length) {
        sendResponse({
          ok: false,
          message: `Completed with issues. ${failures.join(" | ")}`
        });
      } else {
        sendResponse({
          ok: true,
          message: "Bloomberg form filled, added to cart, and proceeded to checkout."
        });
      }
    } catch (err) {
      sendResponse({
        ok: false,
        message: `Unexpected error: ${err.message}`
      });
    }
  })();

  return true;
});
