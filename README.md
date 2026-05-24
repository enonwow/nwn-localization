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
  - Publish path (GitHub API: branch + commit CSV + open PR)
  - Rebuild TLK artifacts from loaded dataset
  - Diff report + JSON/Markdown export

## Run

- Go to app directory: `cd projects/tlk-web`
- Install: `npm install`
- Test: `npm test`
- Dev: `npm run dev`
- Build: `npm run build`

## Publish Target Config (GitHub PR + CSV folder)

- Copy `projects/tlk-web/.env.example` to `projects/tlk-web/.env` and adjust if needed.
- Configure Vite env vars in `projects/tlk-web/.env`:
  - `VITE_GITHUB_PR_REPO=enonwow/nwn-localization-test`
  - `VITE_GITHUB_BASE_BRANCH=main`
  - `VITE_GITHUB_CSV_FOLDER=csv-latest`
  - `VITE_GITHUB_PUBLIC_TOKEN=` (optional local fallback only)
- Alternative repo input is also supported:
  - `VITE_GITHUB_PR_REPO_URL=https://github.com/enonwow/nwn-localization/tree/main/csv-latest`
- If env vars are missing, defaults are:
  - repo: `enonwow/nwn-localization-test`
  - branch: `main`
  - folder: `csv-latest`
- Required PAT permissions (fine-grained, selected repo only):
  - `Contents: Read and write`
  - `Pull requests: Read and write`
  - `Metadata: Read-only`
- GitHub Pages runtime:
  - user pastes PAT in Publish screen (`GitHub PAT (test)`), stored only in browser session.
  - workflow file: `.github/workflows/deploy-pages.yml`
