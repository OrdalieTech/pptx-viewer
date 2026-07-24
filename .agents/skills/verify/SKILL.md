---
name: verify
description: Verify pptx-viewer viewer or editor changes in the live React, Vue, Angular, Vanilla, and Svelte demos. Use after changing a UI binding, shared rendering logic, core behavior visible in a viewer, or demo integration, especially when the user asks to test or validate changes in a browser.
---

# Verify pptx-viewer changes

Validate the smallest set of demos affected by the change. Core or shared
rendering changes generally affect all five bindings; binding-only changes
normally require only that binding's demo.

## Prepare built dependencies

The demos do not resolve every package from source:

| Package used by demo | React  | Vue    | Angular  | Vanilla | Svelte |
| -------------------- | ------ | ------ | -------- | ------- | ------ |
| binding              | source | source | dist     | source  | source |
| core                 | source | source | source   | source  | source |
| shared               | dist   | source | vendored | source  | source |
| locales              | source | source | dist     | source  | source |
| tools                | dist   | dist   | dist     | dist    | dist   |

Before launching:

- Build `packages/angular` after Angular or shared changes.
- Build `packages/shared` after adding or changing an export used by React.
- Build `packages/tools` after tools changes or when every demo fails with a
  browser compatibility error involving Node's `path` module.

## Launch

Check whether the relevant server is already listening before starting another:

```bash
lsof -nP -iTCP:4173-4177 -sTCP:LISTEN
```

Run the needed command from the repository root:

- React: `bun run demo` on port 4173
- Angular: `bun run demo:angular` on port 4174
- Vue: `bun run demo:vue` on port 4175
- Vanilla: `bun run demo:vanilla` on port 4176
- Svelte: `bun run demo:svelte` on port 4177

Reuse an existing server when it serves current code. Do not terminate another
session's server unless stale behavior is confirmed and restarting it is needed.

## Exercise the change

Use the in-app browser control skill for local testing. Use Chrome control only
when existing Chrome state is specifically needed.

- Create an editable deck from the landing page with "or create a New
  Presentation", or load a fixture from `e2e/fixtures`.
- Reproduce the changed interaction in every selected binding.
- Check visible output and browser console errors.
- For editor changes, save and reload the deck when serialization is relevant.
- Capture screenshots when visual comparison materially helps verification.

Run the narrowest package tests and typecheck that cover the change in addition
to live verification.

## Troubleshooting

- If behavior contradicts source, restart a stale Vite server.
- If a stack trace points inside a framework runtime, clear that demo's
  `node_modules/.vite` cache and restart before diagnosing application code.
- Browser screenshots can time out immediately after an interaction; retry
  once after the UI settles.
- The working tree and running servers may be shared with other sessions.
