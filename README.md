# nwn-localization

## Documentation Rules

- Documentation is maintained only in this `README.md`.
- Do not create or reference additional documentation folders/files.
- Aurora First: TLK behavior assumptions come from Aurora reverse-engineering first.
- Development mode: TDD (`red -> green -> refactor`) for new logic and fixes.

## Required Controller

- Implement a controller endpoint: `POST /api/tlk/convert`.
- The endpoint must accept an XLSX file upload (request body via `multipart/form-data`).
- The endpoint must validate the XLSX format.
- If valid, the endpoint must convert XLSX to TLK and return TLK in the response.

## Current Frontend Status (MVP)

- Frontend stack migrated to:
  - React + TypeScript + Vite
  - AG Grid Community (MIT)
- Project structure:
  - `projects/tlk-web/` - web app (current working project)
- Main app entry: `projects/tlk-web/src/features/workflow/LocalizationWorkflow.tsx`.
- Core domain modules:
  - `projects/tlk-web/src/lib/tlk.ts`
  - `projects/tlk-web/src/lib/xlsx.ts`
  - `projects/tlk-web/src/lib/validation.ts`
  - `projects/tlk-web/src/lib/diff.ts`
- Unit tests (Vitest): `projects/tlk-web/src/lib/__tests__/core.test.ts`.
- Current browser-side flow:
  - Import TLK (locale packs + fallback `dialogf -> dialog`)
  - Import CSV (with legacy XLSX compatibility)
  - Edit grid with paging and undo/redo
  - Validate
  - Export CSV
  - Publish path (PR mock + approved CSV import)
  - Rebuild TLK artifacts from loaded dataset
  - Diff report + JSON/Markdown export

## Run

- Go to app directory: `cd projects/tlk-web`
- Install: `npm install`
- Test: `npm test`
- Dev: `npm run dev`
- Build: `npm run build`
