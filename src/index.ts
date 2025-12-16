import readXlsxFile from "read-excel-file/node";
import playwright, { Page } from "playwright";
const { chromium } = require("playwright-extra");
import * as fs from "node:fs/promises";
import path from 'path';
import { PDFParse } from 'pdf-parse';


const excelFile = "./mockExcel.xlsx"

interface ExcelItem {
  brand: string;
  article: string;
  description: string;
}

interface ResultItem {
  "Номер позиции из файла": number;
  Наименование: string;
  Описание: string;
  Характеристики: Record<string, string>;
}

type SpecDict = Record<string, string>;

async function excelReader(file: string): Promise<ExcelItem[]> {
  try {
    const rows = await readXlsxFile(file);
    const items = rows.slice(1).map((row) => ({
      brand: String(row[0]),
      article: String(row[1]),
      description: String(row[2]),
    }));

    return items;
  } catch (err) {
    console.log("error while reading excel", err);
    return [];
  }
}

async function search(item: string): Promise<string[]> {
  const browser = await playwright["chromium"].launch({
    headless: false,
    slowMo: 150,
  });
  const context = await browser.newContext();
  const page = await context.newPage();

  try {
    const query = encodeURIComponent(item + " -site:ozon.ru");
    const url = `https://html.duckduckgo.com/html/?q=${query}&kl=ru-ru`;

    await page.goto(url);

    const results = await page.$$eval(".result", (elements) => {
      return elements
        .slice(0, 5)
        .map(
          (el) => el.querySelector(".result__url")?.textContent?.trim() || "",
        )
        .filter((url) => url);
    });

    return results;
  } catch (error) {
    console.error("Ошибка при поиске!!", error);
    return [];
  } finally {
    await context.close();
    await browser.close();
  }
}

function saveJSON(results: ResultItem[]) {
  fs.writeFile("./src/results.json", JSON.stringify(results, null, 2), {
    flag: "a",
  });
}

const USER_AGENTS = [
  "Mozilla/5.0 (Windows NT 10.0; Win64; x64) AppleWebKit/537.36 (KHTML, like Gecko) Chrome/120.0.0.0 Safari/537.36",
  "Mozilla/5.0 (Macintosh; Intel Mac OS X 10_15_7) AppleWebKit/537.36 (KHTML, like Gecko) Chrome/119.0.0.0 Safari/537.36",
  "Mozilla/5.0 (X11; Linux x86_64) AppleWebKit/537.36 (KHTML, like Gecko) Chrome/120.0.0.0 Safari/537.36",
];

async function parseSite(sites: string[], iter: number): Promise<string[]> {
  const browser = await chromium.launch({
    headless: false,
    args: ["--disable-blink-features=AutomationControlled"],
  });

  try {
    for (const site of sites) {
      const context = await browser.newContext({
        userAgent: USER_AGENTS[Math.floor(Math.random() * USER_AGENTS.length)],
        viewport: { width: 1920, height: 1080 },
        locale: "ru-RU",
        extraHTTPHeaders: {
          "Accept-Language": "ru-RU,ru;q=0.9,en;q=0.8",
        },
      });

      await context.route("**/*", (route: any) => {
        const resourceType = route.request().resourceType();
        if (["font", "media"].includes(resourceType)) {
          route.abort();
        } else {
          route.continue();
        }
      });

      const page = await context.newPage();
      const url = site.startsWith("http") ? site : `https://${site}`;

      if(url.endsWith('.pdf')) {
        console.log("парсим pdf")
        const parser = new PDFParse({ url: url });
        const result = await parser.getText();
        await parser.destroy();
        const start = result.text.search(/характеристики/i);
        const tail = result.text.slice(start);
        const end = tail.search("\n\n");
        const specs = result.text.slice(start, start+end);
        console.log("specs", specs)
        if(specs.length > 10) {
          return [specs];
        }
        continue;
      }

      console.log(`[START] Пробуем: ${url}`);

      try {
        const response = await page.goto(url, {
          timeout: 25000,
          waitUntil: "domcontentloaded",
        });

        if (!response || response.status() >= 400) {
          console.warn(`[SKIP] HTTP ${response?.status()} на ${url}`);
          await context.close();
          continue;
        }

        await page.waitForTimeout(1000);
        await expandSpecsContent(page);

        const specs = await extractSpecsFromPage(page);

        if (specs.length > 0) {
          console.log(
            `[SUCCESS] Успешно извлечено ${specs.length} строк с ${url}`,
          );
          await downloadMainPicture(page, iter);
          await context.close();
          return specs;
        } else {
          console.log(`[FAIL] Характеристики не найдены на ${url}`);
        }
      } catch (err) {
        console.error(
          `[ERROR] Ошибка на ${url}:`,
          err instanceof Error ? err.message : err,
        );
      } finally {
        await context.close();
      }
    }

    return ["Не удалось извлечь характеристики ни с одного сайта"];
  } finally {
    await browser.close();
  }
}

const downloadMainPicture = async (page: Page, number: number) => {
  const imageUrl = await page.evaluate(() => {
    window.scrollTo(0,0);
    const images = Array.from(document.querySelectorAll("img"));
    const valid = [];

    for (let img of images) {
      if (img.naturalHeight > 300 && img.naturalWidth > 300) {
        valid.push({
          src: img.src,
          width: img.naturalWidth,
          height: img.naturalHeight,
        });
      }
    }

    const neededImg = valid.sort((a, b) => b.width - a.width)[0];
    return neededImg ? neededImg.src : null;
  });
  // console.log("img url", imageUrl)
  await page.waitForTimeout(1000); 
  if (imageUrl) {
    try {
      console.log(`[DOWNLOAD] Скачиваем картинку: ${imageUrl}`);
      const response = await page.request.get(imageUrl);
      if (response.ok()) {
        const urlObj = new URL(imageUrl);
        const cleanPath = urlObj.pathname;                    
        const ext = cleanPath.split('.').pop() || 'jpg';       
        
        const buffer = await response.body();
        const dir = `images/${number}`;

        await fs.mkdir(dir, { recursive: true }); 
        await fs.writeFile(path.join(dir, `${number}.${ext}`), buffer);  
        await page.screenshot({                       
          path: path.join(dir, `${number}_2.png`)
        }); //еще один скриншот на всякий но захватывает всю страницу

        console.log(`[DOWNLOAD] Картинка ${number}.jpg сохранена`);

      } else {
        console.warn(
          `[DOWNLOAD] Не удалось скачать картинку, статус: ${response.status()}`,
        );
      }
    } catch (e) {
      console.error(`[DOWNLOAD] Ошибка при скачивании: ${e}`);
    }
  } else {
    console.warn(`[DOWNLOAD] Подходящая картинка не найдена`);
  }
};

async function expandSpecsContent(page: Page) {
  const expandKeywords = [
    "все характеристики",
    "полные характеристики",
    "показать характеристики",
    "технические характеристики",
    "показать все",
    "характеристики",
    "развернуть",
    "подробнее",
    "смотреть все",
    "читать далее",
  ];

  for (const keyword of expandKeywords) {
    try {
      const elements = await page
        .locator(
          `button:has-text("${keyword}"), a:has-text("${keyword}"), div[role="button"]:has-text("${keyword}")`,
        )
        .all();

      for (const el of elements) {
        if (await el.isVisible()) {
          console.log(`[ACTION] Кликаем по "${keyword}"`);
          await el.click({ timeout: 2000, force: true }).catch(() => {});
          await page.waitForTimeout(1500);
          return;
        }
      }
    } catch (e) {
      continue;
    }
  }

  const tabSelectors = [
    '[role="tab"]:has-text("Характеристики")',
    '[role="tab"]:has-text("Описание")',
    '.tab:has-text("Характеристики")',
    'a[href*="specs"]',
    'a[href*="характеристик"]',
  ];

  for (const selector of tabSelectors) {
    try {
      const tab = page.locator(selector).first();
      if (await tab.isVisible()) {
        console.log(`[ACTION] Кликаем по вкладке: ${selector}`);
        await tab.click({ timeout: 2000 }).catch(() => {});
        await page.waitForTimeout(1500);
        return;
      }
    } catch (e) {
      continue;
    }
  }

  const accordionSelectors = [
    '.accordion:has-text("Характеристики")',
    '[class*="accordion"]:has-text("Характеристики")',
    'details summary:has-text("Характеристики")',
  ];

  for (const selector of accordionSelectors) {
    try {
      const accordion = page.locator(selector).first();
      if (await accordion.isVisible()) {
        console.log(`[ACTION] Раскрываем аккордеон: ${selector}`);
        await accordion.click({ timeout: 2000 }).catch(() => {});
        await page.waitForTimeout(1000);
        return;
      }
    } catch (e) {
      continue;
    }
  }
}

async function extractSpecsFromPage(page: Page): Promise<string[]> {
  return await page.evaluate(() => {
    const cleanText = (text: string | null | undefined): string => {
      if (!text) return "";
      return text
        .replace(/\s+/g, " ")
        .replace(/[\u200B-\u200D\uFEFF]/g, "")
        .trim();
    };

    document
      .querySelectorAll(
        'script, style, nav, footer, header, .ad, .banner, iframe, noscript, [class*="cookie"], [class*="popup"]',
      )
      .forEach((el) => el.remove());

    const candidates: { element: Element; score: number; text: string[] }[] =
      [];

    
    const possibleContainers = document.querySelectorAll(
      `table, dl, ul.specs, ul[class*="spec"], ul[class*="characteristic"], ul[class*="attribute"],
       div[class*="specs"], div[class*="spec"], div[class*="характерист"],
       div[class*="attributes"], div[class*="properties"], div[class*="params"],
       section[class*="spec"], section[class*="detail"], section[class*="info"],
       .product-info, .product-details, .product-specs, .product-attributes,
       .characteristics, .tech-specs, .technical-info,
       [id*="spec"], [id*="characteristic"], [id*="detail"]`,
    );

    possibleContainers.forEach((container) => {
      let score = 0;
      const lines: string[] = [];
      const textContent = container.textContent?.toLowerCase() || "";

      if (
        textContent.includes("характерист") ||
        textContent.includes("свойства")
      )
        score += 10;
      if (
        textContent.includes("specification") ||
        textContent.includes("features")
      )
        score += 10;
      if (textContent.includes("вес") || textContent.includes("weight"))
        score += 3;
      if (
        textContent.includes("размер") ||
        textContent.includes("габарит") ||
        textContent.includes("size")
      )
        score += 3;
      if (textContent.includes("цвет") || textContent.includes("color"))
        score += 2;
      if (textContent.includes("материал") || textContent.includes("material"))
        score += 2;
      if (textContent.includes("гарантия") || textContent.includes("warranty"))
        score += 2;
      if (
        textContent.includes("производитель") ||
        textContent.includes("бренд")
      )
        score += 2;

      if (container.tagName === "TABLE") {
        const rows = container.querySelectorAll("tr");
        rows.forEach((row) => {
          const cells = row.querySelectorAll("td, th");
          if (cells.length >= 2) {
            const key = cleanText(cells[0].textContent);
            const val = cleanText(cells[1].textContent);
            if (
              key &&
              val &&
              key.length < 150 &&
              val.length < 500 &&
              key !== val
            ) {
              lines.push(`${key}: ${val}`);
              score += 5;
            }
          }
        });
      } else if (container.tagName === "DL") {
        const dts = Array.from(container.querySelectorAll("dt"));
        const dds = Array.from(container.querySelectorAll("dd"));

        dts.forEach((dt, i) => {
          const key = cleanText(dt.textContent);
          const val = cleanText(dds[i]?.textContent);
          if (key && val && key !== val) {
            lines.push(`${key}: ${val}`);
            score += 4;
          }
        });
      } else if (container.tagName === "UL" || container.tagName === "OL") {
        const items = container.querySelectorAll("li");
        items.forEach((item) => {
          const keyEl = item.querySelector(
            '[class*="name"], [class*="key"], [class*="label"], [class*="title"]',
          );
          const valEl = item.querySelector(
            '[class*="value"], [class*="val"], [class*="data"]',
          );

          if (keyEl && valEl) {
            const key = cleanText(keyEl.textContent);
            const val = cleanText(valEl.textContent);
            if (key && val && key !== val) {
              lines.push(`${key}: ${val}`);
              score += 3;
            }
          } else {
            const text = cleanText(item.textContent);
            const match = text.match(
              /^([\wа-яА-Я\s().,%№+-]+?)\s*(?:[:\.—\-]{1,}|\s{3,})\s*(.{1,300})$/, // /^([\wа-яА-Я\s().,%№+-]+?)\s*[:\.—\-]{1,}|\s{3,}\s*(.{1,300})$/
            );
            if (match && match[1] && match[2]) {
              lines.push(`${match[1].trim()}: ${match[2].trim()}`);
              score += 2;
            }
          }
        });
      } else {
        const pairs = container.querySelectorAll(
          '[class*="row"], [class*="item"], [class*="line"], [class*="param"], [class*="attr"]',
        );

        pairs.forEach((pair) => {
          const keyEl = pair.querySelector(
            '[class*="name"], [class*="key"], [class*="label"], [class*="title"], [class*="prop"]',
          );
          const valEl = pair.querySelector(
            '[class*="value"], [class*="val"], [class*="data"], [class*="content"]',
          );

          if (keyEl && valEl) {
            const key = cleanText(keyEl.textContent);
            const val = cleanText(valEl.textContent);
            if (
              key &&
              val &&
              key.length < 150 &&
              val.length < 500 &&
              key !== val
            ) {
              lines.push(`${key}: ${val}`);
              score += 3;
            }
          }
        });

        if (lines.length === 0) {
          const rawText =
            container instanceof HTMLElement
              ? container.innerText
              : container.textContent || "";
          const splitLines = rawText.split("\n");

          splitLines.forEach((line) => {
            const cleaned = cleanText(line);
            const match = cleaned.match(
              /^([\wа-яА-Я\s().,%№+-]{2,50}?)\s*([:\.—\-]{1,}|\s{3,})\s*(.{1,300})$/,
            );

            if (
              match &&
              match[1] &&
              match[3] &&
              !cleaned.includes("{") &&
              !cleaned.includes("function") &&
              !cleaned.includes("var ") &&
              !cleaned.includes("return") &&
              match[1].trim() !== match[3].trim() &&
              match[3].length > 1
            ) {
              lines.push(`${match[1].trim()}: ${match[3].trim()}`);
              score += 1;
            }
          });
        }
      }

      if (lines.length > 2) {
        candidates.push({ element: container, score, text: lines });
      }
    });

    candidates.sort((a, b) => b.score - a.score);

    return candidates.length > 0 ? candidates[0].text : [];
  });
}

const keyValueParse = (specs: string[]) => {
  const obj: Record<string, string> = {};

  for(let spec of specs) {
    let indexDots = spec.indexOf(":")
    let key = spec.substring(0, indexDots);
    let value = spec.substring(indexDots+1, spec.length);

    obj[key] = value;
  }

  return obj;
}


async function main() {
  let items: ExcelItem[] = await excelReader(excelFile);
  const results: ResultItem[] = [];

  for (let i = 0; i < items.length; i++) {
    try {
      let item = items[i];
      let itemRow = item.brand + " " + item.article + " " + item.description;

      const searchResult = await search(itemRow);
      console.log("search result", searchResult);

      const parseResult = await parseSite(searchResult, i);
      console.log("parse result", parseResult);
      const specs = keyValueParse(parseResult);

      const resultItem: ResultItem = {
        "Номер позиции из файла": i,
        Наименование: item.brand + " " + item.article,
        Описание: item.description,
        Характеристики: specs ?? "не найдено",
      };
      results.push(resultItem);

    } catch (error) {
      console.log("что-то пошло не так", error);
    }
  }
  saveJSON(results);
}

main().catch(console.error);