# DEXY CD Converter platform foundation

This document describes the first database-backed version of DEXY CD Converter.
The current production app remains a browser-side converter. The platform phase
adds storage, authentication, audit history, and traceability without changing
the existing conversion rules.

## Target architecture

- Netlify hosts the React application.
- Supabase hosts authentication, PostgreSQL data, and file storage.
- The browser keeps doing Excel parsing and conversion.
- Conversion metadata is saved after each successful conversion.
- Original files and generated files can later be uploaded to Supabase Storage.

## First data to store

- Organizations using the converter.
- Users and their roles inside each organization.
- Conversion templates and rules by company.
- Conversion batches.
- Imported and generated files.
- Certified invoices, including DGI code, currency, exchange rate, and totals.
- Credit note links to their original certified invoices.
- Processing events and errors for audit.

## Why this matters for USD credit notes

For USD credit notes, the credit note must reuse the exchange rate of the
original certified invoice. Today, if DEXY does not export that rate, the app can
infer it from certified CDF totals as a fallback. Long term, the platform should
store the original rate when the sales invoice is converted or certified.

## Branching rule

- `main` stays stable and production-ready.
- `develop` receives platform work.
- Netlify preview deploys can be used for validation.
- Merge to `main` only after business tests are approved.
