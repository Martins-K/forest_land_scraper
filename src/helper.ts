import { Page } from "playwright";

export const getDetailsText = async (page: Page, label: string): Promise<string> => {
  const labelLocator = page.locator("td.ads_opt_name", { hasText: label });

  if ((await labelLocator.count()) > 0) {
    const valueLocator = labelLocator.locator("xpath=following-sibling::td[1]");
    try {
      const rawText = await valueLocator.innerText();
      let cleanText = rawText.trim();

      cleanText = cleanText.replace(/\s*\[\s*karte\s*\]/gi, "").trim();

      if (cleanText === "-") return "";

      return cleanText;
    } catch {
      return "";
    }
  }

  return "";
};

export const getBrand = async (page: Page): Promise<string> => {
  const labelLocator = page.locator("td.ads_opt_name", { hasText: "Marka" });

  if ((await labelLocator.count()) > 0) {
    const valueLocator = labelLocator.locator("xpath=following-sibling::td[1]");

    try {
      const rawText = await valueLocator.innerText();
      let cleanText = rawText.trim();

      if (cleanText === "-" || cleanText === "") return "";

      return cleanText;
    } catch {
      return "";
    }
  }

  return "";
};

export const getDescription = async (page: Page): Promise<string> => {
  try {
    const descriptionElement = page.locator("#msg_div_msg");

    if ((await descriptionElement.count()) > 0) {
      // Get the text content before the first table
      const fullText = await descriptionElement.evaluate((el) => {
        // Get all text nodes before the first table
        const walker = document.createTreeWalker(el, NodeFilter.SHOW_TEXT, {
          acceptNode: (node) => {
            // Stop at table elements
            let parent = node.parentElement;
            while (parent && parent !== el) {
              if (parent.tagName === "TABLE") {
                return NodeFilter.FILTER_REJECT;
              }
              parent = parent.parentElement;
            }
            return NodeFilter.FILTER_ACCEPT;
          },
        });

        let textContent = "";
        let node;
        while ((node = walker.nextNode())) {
          textContent += node.textContent;
        }

        return textContent;
      });
      return fullText.replace(/\s+/g, " ").trim();
    }
  } catch (error) {
    console.log(`Error extracting description: ${error}`);
  }

  return "";
};
