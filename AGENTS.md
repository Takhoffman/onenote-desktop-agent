# Agent Instructions

This repository intentionally contains two copies of the `onenote-desktop` skill.

The source of truth is:

```text
skills/onenote-desktop/
```

The Codex marketplace package copy is:

```text
plugins/onenote-desktop/skills/onenote-desktop/
```

Keep the root skill authoritative. After editing anything under `skills/onenote-desktop/`, run:

```powershell
powershell.exe -NoProfile -ExecutionPolicy Bypass -File .\scripts\sync-plugin-package.ps1
```

Before committing or releasing, verify the package copy has not drifted:

```powershell
powershell.exe -NoProfile -ExecutionPolicy Bypass -File .\scripts\check-plugin-package.ps1
```

Do not edit the packaged copy directly unless you are deliberately changing packaging behavior. If the packaged copy needs content changes, make them in the root skill and sync.

The duplication exists because the root plugin layout supports direct/local installs, while `.agents/plugins/marketplace.json` points Codex marketplace installs at `plugins/onenote-desktop/`.
