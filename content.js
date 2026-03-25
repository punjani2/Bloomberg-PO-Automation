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

function canonicalLabel(text) {
  return normalizeText(text).replace(/[:*]/g, "").trim();
}

function findRowByLabel(labelText) {
  const target = canonicalLabel(labelText);
  const labels = Array.from(document.querySelectorAll("label"));
  for (const label of labels) {
    const labelTextNormalized = canonicalLabel(label.textContent);
    if (labelTextNormalized === target || labelTextNormalized.startsWith(target)) {
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

function clickElementRobust(el) {
  if (!el) return;
  el.scrollIntoView({ block: "center", inline: "center" });
  el.dispatchEvent(new MouseEvent("mousedown", { bubbles: true }));
  el.dispatchEvent(new MouseEvent("mouseup", { bubbles: true }));
  el.dispatchEvent(new MouseEvent("click", { bubbles: true }));
}

function getContactRowSelectedText() {
  const contactRow = findRowByLabel("Contact:");
  if (!contactRow) return "";

  const valueEl = contactRow.querySelector("input.w-chInput, input[type='text'], a");
  if (!valueEl) return normalizeText(contactRow.textContent);

  const raw = valueEl.tagName === "INPUT" ? valueEl.value : valueEl.textContent;
  return normalizeText(raw);
}

function getChooserRowValueText(label) {
  const row = findRowByLabel(label);
  if (!row) return "";

  const valueEl = row.querySelector("input.w-chInput, input[type='text'], a");
  if (!valueEl) return normalizeText(row.textContent);

  const raw = valueEl.tagName === "INPUT" ? valueEl.value : valueEl.textContent;
  return normalizeText(raw);
}

async function selectVendorViaSearchMore() {
  logStep("Selecting Vendor");

  const vendorRow = findRowByLabel("Vendor:");
  if (!vendorRow) {
    throw new Error("Row not found for Vendor:");
  }

  const vendorInput = vendorRow.querySelector("input.w-chInput");
  if (!vendorInput) throw new Error("Vendor chooser input not found");

  setInputValue(vendorInput, "Bloomberg Finance");
  vendorInput.dispatchEvent(new Event("blur", { bubbles: true }));
  await sleep(400);

  const targetId = normalizeText("0002190112");
  const targetName = normalizeText("BLOOMBERG FINANCE L P");

  // First attempt: inline dropdown option under Vendor field (matches user's screenshot flow).
  try {
    const wrapper = vendorRow.querySelector(".w-chWrapper");
    const menuId = wrapper?.querySelector(".w-chWrap")?.getAttribute("awmenuid");
    if (!menuId) throw new Error("Vendor inline menu id not found");

    const arrow =
      vendorRow.querySelector(".w-chPullDownDiv") ||
      vendorRow.querySelector(".w-chWrapRight a") ||
      vendorRow.querySelector(".w-chWrapRight");
    if (!arrow) throw new Error("Vendor inline dropdown arrow not found");

    clickElementRobust(arrow);
    await sleep(300);

    const menu = await waitFor(() => document.getElementById(menuId), 4000, 200);
    if (!menu) throw new Error("Vendor inline dropdown menu not found");

    const option = await waitFor(() => {
      const opts = Array.from(menu.querySelectorAll("[role='option'], .w-pmi-item"));
      return opts.find(opt => {
        const text = normalizeText(opt.textContent);
        return text.includes(targetId) && text.includes(targetName);
      });
    }, 5000, 200);

    if (!option) throw new Error("Vendor inline dropdown option not found");

    clickElementRobust(option);
    await sleep(700);
  } catch (inlineErr) {
    logStep(`Vendor inline dropdown fallback to Search More: ${inlineErr.message}`);

    const searchMoreControl = Array.from(vendorRow.querySelectorAll("button, a, [role='button'], span")).find(
      el => normalizeText(el.textContent).includes("search more")
    );

    if (!searchMoreControl) {
      throw new Error(`Vendor inline dropdown failed and Search More not found: ${inlineErr.message}`);
    }

    searchMoreControl.click();

    const popup = await waitFor(() => {
      const dialogs = Array.from(document.querySelectorAll("[role='dialog'], .w-window, .w-dlg"));
      return dialogs.find(el => isElementVisible(el) && normalizeText(el.textContent).includes("choose value for vendor"));
    }, 7000, 200);

    if (!popup) {
      throw new Error('Vendor popup "Choose Value for Vendor" did not appear');
    }

    const popupSearch = Array.from(popup.querySelectorAll("button, a, [role='button']")).find(
      el => normalizeText(el.textContent) === "search"
    );
    if (popupSearch) {
      popupSearch.click();
      await sleep(700);
    }

    const vendorResultRow = await waitFor(() => {
      const rows = Array.from(popup.querySelectorAll("tr, [role='row'], .w-tbl-row"));
      return rows.find(row => {
        const text = normalizeText(row.textContent);
        return text.includes(targetId) && text.includes(targetName);
      });
    }, 8000, 200);

    if (!vendorResultRow) {
      throw new Error("Vendor row 0002190112 BLOOMBERG FINANCE L P not found in popup");
    }

    const clickVendorSelectButton = async () => {
      const rowSelect = Array.from(vendorResultRow.querySelectorAll("button, a, [role='button']")).find(
        el => normalizeText(el.textContent) === "select"
      );
      if (rowSelect) {
        clickElementRobust(rowSelect);
        await sleep(700);
        return true;
      }

      // Fallback: click the second visible "Select" in popup (first is usually No Preference).
      const popupSelectButtons = Array.from(popup.querySelectorAll("button, a, [role='button']")).filter(
        el => normalizeText(el.textContent) === "select" && isElementVisible(el)
      );
      if (popupSelectButtons.length >= 2) {
        clickElementRobust(popupSelectButtons[1]);
        await sleep(700);
        return true;
      }

      return false;
    };

    const selectClicked = await clickVendorSelectButton();
    if (!selectClicked) {
      throw new Error('Vendor row "Select" button not found/clickable');
    }

    const doneBtn = Array.from(popup.querySelectorAll("button, a, [role='button']")).find(
      el => normalizeText(el.textContent) === "done"
    );
    if (doneBtn && isElementVisible(doneBtn)) {
      doneBtn.click();
      await sleep(600);
    }
  }

  const selectedVendor = getChooserRowValueText("Vendor:");
  if (!selectedVendor.includes(targetId) || !selectedVendor.includes(targetName)) {
    throw new Error("Vendor did not populate with 0002190112 BLOOMBERG FINANCE L P");
  }
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

  const ready = await waitFor(() => {
    const row = findRowByLabel("Account Type:");
    if (!row) return null;

    const combo = row.querySelector("[role='combobox'].w-dropdown");
    if (!combo) return null;

    clickElementRobust(combo);

    const comboId = combo.getAttribute("id");
    const itemsContainer = comboId ? document.getElementById(`Items_${comboId}`) : null;
    if (!itemsContainer) return null;

    const options = Array.from(
      itemsContainer.querySelectorAll("[role='option'], .w-dropdown-item")
    );

    return options.length > 0 ? { combo, itemsContainer } : null;
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

  clickElementRobust(combo);

  // Fallback: choose last visible option (user indicates WBS element is the last item).
  const fallbackOptions = Array.from(
    itemsContainer.querySelectorAll("[role='option'], .w-dropdown-item")
  ).filter(isElementVisible);
  const lastOption = fallbackOptions[fallbackOptions.length - 1];
  if (lastOption) {
    clickElementRobust(lastOption);
    await sleep(1000);
    const afterText = normalizeText(combo.textContent || combo.innerText || "");
    if (afterText.includes("wbs")) return;
  }

  throw new Error('Option "WBS element" not found in Account Type dropdown (including last-option fallback)');
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

  // Commodity early
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
    await setPlainField("Quantity:", "1");
    results.push("Quantity");
  } catch (e) {
    results.push(`Quantity failed: ${e.message}`);
  }

  try {
    await setPlainField("Price:", "2035.98");
    results.push("Price");
  } catch (e) {
    results.push(`Price failed: ${e.message}`);
  }

  // Vendor via Search More popup
  try {
    await selectVendorViaSearchMore();
    results.push("Vendor");
  } catch (e) {
    results.push(`Vendor failed: ${e.message}`);
    return results;
  }

  try {
    await selectAccountTypeWbs();
    results.push("Account Type");
  } catch (e) {
    results.push(`Account Type failed: ${e.message}`);
    return results;
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

  await sleep(3000);

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
