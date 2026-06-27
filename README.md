# Batch Invoice

Nuxt-based tool for filling KuaTuanTuan shipment templates from courier order Excel files.

## Features

- Upload the KuaTuanTuan unshipped order Excel template.
- Upload the courier order Excel file.
- Prompts for a password when an uploaded workbook is password-protected.
- Parse both workbooks on the server with `exceljs`.
- Match orders with an OpenAI-compatible model, with a local rules fallback.
- Splits one order into multiple rows when it matches multiple tracking numbers.
- Returns a filled `.xlsx` and a success/failure summary.

## Local Development

```bash
pnpm install
pnpm dev
```

Open `http://localhost:3000`.

## Environment Variables

For Kimi Coding:

```bash
OPENAI_API_KEY=your_kimi_key
OPENAI_BASE_URL=https://api.kimi.com/coding/v1
OPENAI_MODEL=kimi-for-coding
```

If no API key is configured, the server uses local rule-based matching.

## Deploy

This project is ready for Vercel. Import the GitHub repository, keep the default Nuxt settings, and add the environment variables above in Vercel Project Settings.
