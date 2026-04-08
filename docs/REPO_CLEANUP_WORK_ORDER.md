# Doc Convert Repo Cleanup Work Order For Claude

Date: 2026-04-08
Scope: Tracked repo state only
Final cleanup requirement: archive this document out of the repo root when the work is complete

## Primary Goal

Keep this context-menu utility repo clean and installation-focused with:
- one root master document
- one obvious install/use/uninstall path
- setup/build/registry scripts grouped cleanly
- no generated installer artifacts living in the root unless intentionally versioned

## Current Issues To Fix

1. The root mixes docs, setup scripts, installers, launcher wrappers, and packaging artifacts.
2. The README is useful, but the repo root still feels like a build/output working directory.
3. The project needs a clearer separation between source scripts, packaging helpers, and release artifacts.
4. The setup/build/install story should be easy to follow from one entry document.

## Work Order

### 1. Keep One Root Master Document
- Keep `README.md` as the root master document.
- Ensure it is the only high-level entrypoint for contributors/users.
- Link to any deeper packaging or developer docs instead of keeping competing top-level notes.

### 2. Clean The Root Layout
- Group related files by purpose where practical:
  - install/uninstall scripts
  - packaging/build scripts
  - core conversion logic
  - docs
- Keep the root minimal and obvious.

### 3. Decide Release Artifact Policy
- Decide whether installer executables or setup outputs belong in the repo.
- If not, remove them from Git and ignore them.
- If yes, document why and where release assets should live.

### 4. Reconcile Build And Setup Paths
- Ensure `build.cmd`, `setup.cmd`, `setup.ps1`, `install.ps1`, `uninstall.ps1`, and `launcher.vbs` reflect one coherent supported workflow.
- Move historical or redundant script paths into an archive/tools folder if needed.

### 5. Add Maintenance Guardrails
- Add a lightweight repo policy covering:
  - one root master document
  - release artifacts not tracked by default
  - install/build scripts kept current and intentional
  - root remains clean and installation-focused

## Acceptance Criteria
- A user can tell how to install and use the tool from the root immediately.
- Packaging/install scripts are organized and current.
- The root no longer looks like a mixed source/output directory.

## Final Deliverable
- short cleanup report with files moved, removed, rewritten, archived, and any unresolved release-artifact decisions

## Archive Instruction
- When done, move this file out of the repo root into an archive/docs-history location.
