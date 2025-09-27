import { chromium } from "playwright";
import * as fs from "fs";
import ExcelJS from "exceljs";
import dotenv from "dotenv";
import { getDetailsText } from "./helper";
import { sendEmail } from "./emailServiceLands";

dotenv.config();

// Helper function to parse date string to Date object
function parseDate(dateStr: string): Date {
  const [datePart, timePart] = dateStr.split(" ");
  const [day, month, year] = datePart.split(".").map(Number);
  const [hours, minutes] = timePart.split(":").map(Number);
  return new Date(year, month - 1, day, hours, minutes);
}

// Helper function to determine cutoff hours based on day of week
function getCutoffHours(): number {
  const now = new Date();
  const dayOfWeek = now.getDay(); // 0 = Sunday, 1 = Monday, ..., 6 = Saturday

  // If it's Monday (day 1), use 73 hours (3 days) to cover the weekend
  // For other weekdays (Tuesday-Friday), use 25 hours (1 day + 1 hour buffer)
  // For weekends, use 25 hours as well (though the cron only runs Mon-Fri)
  return dayOfWeek === 1 ? 73 : 25;
}

// Normalise various Excel cell value shapes into a plain string
function cellValueToString(v: any): string {
  if (v === null || v === undefined) return "";
  if (typeof v === "string") return v.trim();
  if (typeof v === "number") return String(v);
  if (typeof v === "boolean") return v ? "true" : "false";
  if (v instanceof Date) return v.toISOString();
  if (typeof v === "object") {
    // Hyperlink objects: { text: '...', hyperlink: 'https://...' }
    if ("hyperlink" in v && v.hyperlink) return String(v.hyperlink).trim();
    if ("text" in v && typeof v.text === "string") return v.text.trim();
    // Rich text: { richText: [{text: 'a'}, ...] }
    if ("richText" in v && Array.isArray(v.richText))
      return v.richText
        .map((rt: any) => rt.text || "")
        .join("")
        .trim();
    // Formula result or other objects
    if ("result" in v) return String(v.result || "").trim();
    try {
      return JSON.stringify(v);
    } catch {
      return String(v);
    }
  }
  return String(v);
}

// Data type
interface LandData {
  link: string;
  price: string;
  districtText?: string;
  areaText?: string;
  cadastreText?: string;
  date: string;
  parsedDate?: Date; // Add parsed date for Excel formatting
}

const FILE_NAME = "lands-scraped.xlsx";

async function run() {
  const browser = await chromium.launch({ headless: true });
  const page = await browser.newPage();

  const data: LandData[] = [];
  const scrapedLinksInThisRun = new Set<string>();

  const now = new Date();
  const cutoffHours = getCutoffHours();
  // Scraping only items newer than cutoff hours
  const cutoffDate = new Date(now.getTime() - cutoffHours * 60 * 60 * 1000);
  console.log(`Scraping only items newer than: ${cutoffDate.toISOString()} (${cutoffHours} hours)`);

  const urls = [
    "https://www.ss.com/lv/real-estate/plots-and-lands/aizkraukle-and-reg/all/sell/",
    "https://www.ss.com/lv/real-estate/plots-and-lands/aluksne-and-reg/all/sell/",
    "https://www.ss.com/lv/real-estate/plots-and-lands/balvi-and-reg/all/sell/",
    "https://www.ss.com/lv/real-estate/plots-and-lands/bauska-and-reg/all/sell/",
    "https://www.ss.com/lv/real-estate/plots-and-lands/cesis-and-reg/all/sell/",
    "https://www.ss.com/lv/real-estate/plots-and-lands/daugavpils-and-reg/all/sell/",
    "https://www.ss.com/lv/real-estate/plots-and-lands/dobele-and-reg/all/sell/",
    "https://www.ss.com/lv/real-estate/plots-and-lands/gulbene-and-reg/all/sell/",
    "https://www.ss.com/lv/real-estate/plots-and-lands/jekabpils-and-reg/all/sell/",
    "https://www.ss.com/lv/real-estate/plots-and-lands/jelgava-and-reg/all/sell/",
    "https://www.ss.com/lv/real-estate/plots-and-lands/kraslava-and-reg/all/sell/",
    "https://www.ss.com/lv/real-estate/plots-and-lands/kuldiga-and-reg/all/sell/",
    "https://www.ss.com/lv/real-estate/plots-and-lands/liepaja-and-reg/all/sell/",
    "https://www.ss.com/lv/real-estate/plots-and-lands/limbadzi-and-reg/all/sell/",
    "https://www.ss.com/lv/real-estate/plots-and-lands/ludza-and-reg/all/sell/",
    "https://www.ss.com/lv/real-estate/plots-and-lands/madona-and-reg/all/sell/",
    "https://www.ss.com/lv/real-estate/plots-and-lands/ogre-and-reg/all/sell/",
    "https://www.ss.com/lv/real-estate/plots-and-lands/preili-and-reg/all/sell/",
    "https://www.ss.com/lv/real-estate/plots-and-lands/rezekne-and-reg/all/sell/",
    "https://www.ss.com/lv/real-estate/plots-and-lands/saldus-and-reg/all/sell/",
    "https://www.ss.com/lv/real-estate/plots-and-lands/talsi-and-reg/all/sell/",
    "https://www.ss.com/lv/real-estate/plots-and-lands/tukums-and-reg/all/sell/",
    "https://www.ss.com/lv/real-estate/plots-and-lands/valka-and-reg/all/sell/",
    "https://www.ss.com/lv/real-estate/plots-and-lands/valmiera-and-reg/all/sell/",
    "https://www.ss.com/lv/real-estate/plots-and-lands/ventspils-and-reg/all/sell/",
    "https://www.ss.com/lv/real-estate/plots-and-lands/other/all/sell/",
  ];

  for (const url of urls) {
    let shouldStopThisUrl = false;

    try {
      await page.goto(url, { waitUntil: "domcontentloaded" });
    } catch (e) {
      console.warn(`Failed to open ${url}:`, e);
      continue;
    }

    console.log(`Started processing URL: ${url}`);

    let hasNextPage = true;
    let pageCount = 0;

    while (hasNextPage && !shouldStopThisUrl) {
      await page.waitForLoadState("domcontentloaded");
      pageCount++;
      console.log(`Processing page ${pageCount} of ${url}`);

      const listings = await page.$$("td.msg2");
      const cnt = listings.length;
      if (cnt === 0) {
        console.log("No listings found on this page.");
      }

      for (let i = 0; i < cnt; i++) {
        if (shouldStopThisUrl) break;

        // click the i-th listing
        try {
          await listings[i].click();
        } catch (e) {
          // fallback to nth-match locator
          await page.locator(`:nth-match(td.msg2,${i + 1})`).click();
        }

        const link = page.url();

        if (scrapedLinksInThisRun.has(link)) {
          console.log(`Already scraped in this run: ${link}. Stopping this URL.`);
          shouldStopThisUrl = true;
          await page.goBack();
          break;
        }

        scrapedLinksInThisRun.add(link);

        // Scrape fields (wrapped in try/catch for robustness)
        let price = "";
        try {
          price = (await page.locator(".ads_price").innerText()) || "";
        } catch {}

        let districtText = "";
        try {
          districtText = (await getDetailsText(page, "Pilsēta, rajons:")) || "";
        } catch {}

        let areaText = "";
        try {
          areaText = (await getDetailsText(page, "Platība:")) || "";
        } catch {}

        let cadastreText = "";
        try {
          cadastreText = (await getDetailsText(page, "Kadastra numurs:")) || "";
        } catch {}

        let dateStr = "";
        try {
          dateStr = (
            (await page.locator("td.msg_footer", { hasText: "Datums:" }).innerText()) || ""
          ).replace("Datums: ", "");
        } catch {}

        // If date could not be parsed, treat as older and skip
        if (!dateStr) {
          console.log(`No date found for ${link}, skipping item.`);
          await page.goBack();
          continue;
        }

        let currentItemDate: Date;
        try {
          currentItemDate = parseDate(dateStr);
        } catch (e) {
          console.log(`Failed to parse date '${dateStr}' for ${link}, skipping item.`);
          await page.goBack();
          continue;
        }

        // Stop if ad is older than cutoff
        if (currentItemDate < cutoffDate) {
          console.log(`Found ad older than ${cutoffHours}h (${dateStr}). Stopping this URL.`);
          shouldStopThisUrl = true;
          await page.goBack();
          break;
        }

        data.push({
          link,
          price,
          districtText,
          areaText,
          cadastreText,
          date: dateStr,
          parsedDate: currentItemDate, // Store parsed date
        });

        console.log(`Scraped item dated: ${dateStr} | ${link}`);
        await page.goBack();
      }

      if (shouldStopThisUrl) break;

      // Pagination: try to click "Nākamie"
      const nextButton = await page.$('a:has-text("Nākamie")');
      if (nextButton) {
        try {
          await nextButton.click();
          await page.waitForLoadState("domcontentloaded");
        } catch (e) {
          console.log("No more pages or failed to click next. Stopping pagination.");
          hasNextPage = false;
        }
      } else {
        hasNextPage = false;
      }
    }

    console.log(`Finished processing URL: ${url}`);
  }

  await browser.close();

  // Write to Excel (move old New -> Previously added, dedupe, write fresh New)
  await updateExcel(FILE_NAME, data);

  // Optional email - only send if there are new items
  if (
    process.env.EMAIL_USER_LAND &&
    process.env.EMAIL_PASS_LAND &&
    process.env.RECIPIENT_EMAIL_LAND
  ) {
    // Read back the "New" sheet to get only the actually added items
    const workbook = new ExcelJS.Workbook();
    await workbook.xlsx.readFile(FILE_NAME);
    const newSheet = workbook.getWorksheet("New");

    const actuallyNewItems: LandData[] = [];
    if (newSheet && newSheet.rowCount > 1) {
      newSheet.eachRow((row: ExcelJS.Row, rowNumber: number) => {
        if (rowNumber > 1) {
          // Skip header
          const dateValue = row.getCell(6).value;
          let dateStr = "";

          // Handle different date value types
          if (dateValue instanceof Date) {
            dateStr = dateValue
              .toLocaleDateString("lv-LV", {
                day: "2-digit",
                month: "2-digit",
                year: "numeric",
                hour: "2-digit",
                minute: "2-digit",
              })
              .replace(",", "");
          } else {
            dateStr = cellValueToString(dateValue);
          }

          actuallyNewItems.push({
            link: cellValueToString(row.getCell(1).value),
            price: cellValueToString(row.getCell(2).value),
            districtText: cellValueToString(row.getCell(3).value),
            areaText: cellValueToString(row.getCell(4).value),
            cadastreText: cellValueToString(row.getCell(5).value),
            date: dateStr,
          });
        }
      });
    }

    console.log(`Sending email with ${actuallyNewItems.length} actually new items`);
    await sendEmail(actuallyNewItems, FILE_NAME, cutoffHours);
  }
}

async function updateExcel(fileName: string, freshItems: LandData[]): Promise<number> {
  const workbook = new ExcelJS.Workbook();
  const headers = ["link", "price", "districtText", "areaText", "cadastreText", "date"];
  const knownLinks = new Set<string>();

  let hasExistingFile = false;
  if (fs.existsSync(fileName)) {
    try {
      await workbook.xlsx.readFile(fileName);
      hasExistingFile = true;
    } catch (e) {
      console.warn("Failed to read existing workbook, starting a fresh one:", e);
      hasExistingFile = false;
    }
  }

  // Get or create sheets
  let newSheet = workbook.getWorksheet("New");
  let prevSheet =
    workbook.getWorksheet("Previously added") || workbook.addWorksheet("Previously added");

  // Define styles
  const headerStyle = {
    font: { bold: true, size: 8.5 },
    alignment: { vertical: "middle" as const, horizontal: "left" as const },
    fill: {
      type: "pattern" as const,
      pattern: "solid" as const,
      fgColor: { argb: "FFE6E6E6" }, // Light gray background for headers
    },
  };

  const dataStyle = {
    font: { size: 8.5 },
    alignment: { vertical: "middle" as const, horizontal: "left" as const },
  };

  // Define date format for Excel
  const dateFormat = "dd.mm.yyyy hh:mm";

  // Set column widths
  const columnWidths = [
    { width: 70 }, // link
    { width: 17 }, // price
    { width: 13 }, // districtText
    { width: 8 }, // areaText
    { width: 12 }, // cadastreText
    { width: 17 }, // date (increased width for date/time)
  ];

  // Ensure header row exists on "Previously added" sheet with formatting
  const ensureHeader = (sheet: ExcelJS.Worksheet) => {
    const firstRow = sheet.getRow(1);
    const firstCell = firstRow.getCell(1).value;
    if (firstCell === null || firstCell === undefined || String(firstCell).trim() === "") {
      sheet.addRow(headers);
      // Apply header style
      headers.forEach((_, index) => {
        const cell = firstRow.getCell(index + 1);
        cell.font = headerStyle.font;
        cell.alignment = headerStyle.alignment;
        cell.fill = headerStyle.fill;
      });
    }

    // Set column widths
    columnWidths.forEach((col, index) => {
      sheet.getColumn(index + 1).width = col.width;
    });

    // Set date column format
    sheet.getColumn(6).numFmt = dateFormat;

    // Freeze header row
    sheet.views = [
      {
        state: "frozen" as const,
        xSplit: 0,
        ySplit: 1,
        activeCell: "A2",
        showGridLines: true,
      },
    ];
  };

  ensureHeader(prevSheet);

  // If "New" sheet existed, move its content to "Previously added" and gather links
  if (newSheet) {
    newSheet.eachRow((row, rowNumber) => {
      if (rowNumber === 1) return; // skip header
      const linkStr = cellValueToString(row.getCell(1).value);
      if (linkStr) {
        knownLinks.add(linkStr);
        // Add row to prevSheet, ensuring all cells are handled
        const rowData = headers.map((_, index) => {
          const cellValue = row.getCell(index + 1).value;
          // For date column (index 5), try to parse if it's a string
          if (index === 5 && typeof cellValue === "string") {
            try {
              return parseDate(cellValue);
            } catch {
              return cellValue;
            }
          }
          return cellValue;
        });

        const newRow = prevSheet.addRow(rowData);

        // Apply data style to the new row
        newRow.eachCell((cell: ExcelJS.Cell, colNumber: number) => {
          cell.font = dataStyle.font;
          cell.alignment = dataStyle.alignment;

          // Format link column as hyperlink (column 1)
          if (colNumber === 1 && linkStr) {
            cell.value = { text: linkStr, hyperlink: linkStr };
            cell.font = { ...dataStyle.font, color: { argb: "FF0000FF" }, underline: true };
          }

          // Format date column (column 6)
          if (colNumber === 6) {
            cell.numFmt = dateFormat;
          }
        });
      }
    });
    // Remove the old "New" sheet
    workbook.removeWorksheet(newSheet.id);
  }

  // Gather all known links from "Previously added" sheet and apply formatting
  prevSheet.eachRow((row, rowNumber) => {
    if (rowNumber === 1) return; // skip header

    // Apply data style to existing rows
    row.eachCell((cell: ExcelJS.Cell, colNumber: number) => {
      cell.font = dataStyle.font;
      cell.alignment = dataStyle.alignment;

      // Format link column as hyperlink (column 1)
      if (colNumber === 1) {
        const linkStr = cellValueToString(cell.value);
        if (linkStr && linkStr.startsWith("http")) {
          cell.value = { text: linkStr, hyperlink: linkStr };
          cell.font = { ...dataStyle.font, color: { argb: "FF0000FF" }, underline: true };
        }
      }

      // Format date column (column 6)
      if (colNumber === 6) {
        cell.numFmt = dateFormat;
        // If the date is stored as string, convert it to Date object
        if (typeof cell.value === "string") {
          try {
            cell.value = parseDate(cell.value);
          } catch {
            // Keep as string if parsing fails
          }
        }
      }
    });

    const linkStr = cellValueToString(row.getCell(1).value);
    if (linkStr) knownLinks.add(linkStr);
  });

  // Create a new, clean "New" sheet
  newSheet = workbook.addWorksheet("New", { properties: { tabColor: { argb: "FF92D050" } } });

  // Set column widths for New sheet
  columnWidths.forEach((col, index) => {
    newSheet.getColumn(index + 1).width = col.width;
  });

  // Set date column format for New sheet
  newSheet.getColumn(6).numFmt = dateFormat;

  // Add headers to New sheet with formatting
  newSheet.addRow(headers);
  const newHeaderRow = newSheet.getRow(1);
  headers.forEach((_, index) => {
    const cell = newHeaderRow.getCell(index + 1);
    cell.font = headerStyle.font;
    cell.alignment = headerStyle.alignment;
    cell.fill = headerStyle.fill;
  });

  // Freeze header row in New sheet
  newSheet.views = [
    {
      state: "frozen" as const,
      xSplit: 0,
      ySplit: 1,
      activeCell: "A2",
      showGridLines: true,
    },
  ];

  // Insert only fresh, previously unknown items into the new "New" sheet
  let addedCount = 0;
  for (const item of freshItems) {
    if (!item.link || knownLinks.has(item.link)) {
      continue;
    }

    // Use the parsed Date object instead of the string for the date column
    const newRow = newSheet.addRow([
      item.link,
      item.price,
      item.districtText,
      item.areaText,
      item.cadastreText,
      item.parsedDate || item.date, // Use parsedDate if available, fallback to string
    ]);

    // Apply data style to the new row
    newRow.eachCell((cell: ExcelJS.Cell, colNumber: number) => {
      cell.font = dataStyle.font;
      cell.alignment = dataStyle.alignment;

      // Format link column as hyperlink (column 1)
      if (colNumber === 1) {
        cell.value = { text: item.link, hyperlink: item.link };
        cell.font = { ...dataStyle.font, color: { argb: "FF0000FF" }, underline: true };
      }

      // Format date column (column 6) - Set the actual Date object
      if (colNumber === 6) {
        // Ensure we have a Date object
        if (item.parsedDate) {
          cell.value = item.parsedDate;
        } else if (typeof item.date === "string") {
          // Try to parse if we only have the string
          try {
            cell.value = parseDate(item.date);
          } catch {
            cell.value = item.date; // Fallback to string
          }
        }
        cell.numFmt = dateFormat; // Apply the date format
      }
    });

    knownLinks.add(item.link);
    addedCount++;
  }

  // Save workbook
  try {
    await workbook.xlsx.writeFile(fileName);
    console.log(`Excel file saved with formatting: ${fileName}`);
  } catch (e) {
    console.error("Failed to write Excel file:", e);
    throw e;
  }

  const newCount = Math.max(0, newSheet.rowCount - 1);
  const prevCount = Math.max(0, prevSheet.rowCount - 1);
  console.log(
    `Result -> New: ${newCount} rows, Previously added: ${prevCount} rows (added ${addedCount} fresh rows)`
  );

  return addedCount;
}

if (process.env.GITHUB_ACTIONS) {
  // Running in GitHub Actions - just run once
  run().catch((error) => {
    console.error("Scraper failed:", error);
    process.exit(1);
  });
} else {
  // Running locally - use interval
  run();
}
