import https from "node:https";
import { defineConfig } from "vite";
import react from "@vitejs/plugin-react";

const BCC_EXCHANGE_RATE_URL =
  "https://www.bcc.cd/operations-et-marches/domaine-operationnel/operations-de-change/cours-de-change";

export default defineConfig({
  plugins: [
    react(),
    {
      name: "bcc-rate-dev-endpoint",
      configureServer(server) {
        server.middlewares.use(
          "/api/bcc-rate",
          async (_request, response) => {
            try {
              const html = await fetchBccHtml();
              const rate = extractUsdRate(html);
              const publishedAt = extractPublishedAt(html);

              if (rate === null) {
                sendJson(
                  response,
                  {
                    message:
                      "Le taux USD/CDF est introuvable sur la page BCC.",
                  },
                  502,
                );
                return;
              }

              sendJson(response, {
                currency: "USD",
                rate,
                publishedAt,
                sourceUrl: BCC_EXCHANGE_RATE_URL,
              });
            } catch (error) {
              sendJson(
                response,
                {
                  message:
                    error instanceof Error
                      ? error.message
                      : "Lecture du taux BCC impossible.",
                },
                500,
              );
            }
          },
        );
      },
    },
  ],
});

function fetchBccHtml(): Promise<string> {
  return new Promise((resolve, reject) => {
    const request = https.get(
      BCC_EXCHANGE_RATE_URL,
      {
        headers: {
          "user-agent": "DEXY CD Converter/1.0",
        },
        rejectUnauthorized: false,
      },
      (response) => {
        if (!response.statusCode || response.statusCode >= 400) {
          reject(new Error("La BCC n'a pas retourne le taux de change."));
          response.resume();
          return;
        }

        let html = "";
        response.setEncoding("utf8");
        response.on("data", (chunk) => {
          html += chunk;
        });
        response.on("end", () => resolve(html));
      },
    );

    request.setTimeout(15000, () => {
      request.destroy(new Error("La lecture du taux BCC a expire."));
    });
    request.on("error", reject);
  });
}

function extractUsdRate(html: string): number | null {
  const usdRow = Array.from(html.matchAll(/<tr[^>]*>[\s\S]*?<\/tr>/gi))
    .map((match) => match[0])
    .find((row) => /<td>\s*USD\s*<\/td>/i.test(row));

  if (!usdRow) {
    return null;
  }

  const visibleUsdRow = usdRow.replace(/<!--[\s\S]*?-->/g, "");
  const cells = Array.from(visibleUsdRow.matchAll(/<td[^>]*>\s*([^<]+?)\s*<\/td>/gi))
    .map((match) => cleanText(match[1]));

  return parseFrenchNumber(cells[cells.length - 1] ?? "");
}

function extractPublishedAt(html: string): string | null {
  const currentRateBlock = html.match(
    /Cours de change[\s\S]*?<i class="fa fa-calendar"><\/i>\s*([^|<]+)\s*\|\s*<i class="fa fa-clock-o"><\/i>\s*([^<]+)/i,
  );

  if (!currentRateBlock) {
    return null;
  }

  return `${cleanText(currentRateBlock[1])} ${cleanText(currentRateBlock[2])}`;
}

function parseFrenchNumber(value: string): number | null {
  const normalizedValue = cleanText(value)
    .replace(/\s/g, "")
    .replace(",", ".");
  const parsedValue = Number.parseFloat(normalizedValue);
  return Number.isFinite(parsedValue) ? parsedValue : null;
}

function cleanText(value: unknown): string {
  return String(value)
    .replace(/&nbsp;/g, " ")
    .replace(/\s+/g, " ")
    .trim();
}

function sendJson(
  response: { statusCode?: number; setHeader: (name: string, value: string) => void; end: (body: string) => void },
  payload: unknown,
  statusCode = 200,
): void {
  response.statusCode = statusCode;
  response.setHeader("cache-control", "no-store");
  response.setHeader("content-type", "application/json; charset=utf-8");
  response.end(JSON.stringify(payload));
}
