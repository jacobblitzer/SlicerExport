# Agent Guide — SlicerExport

> **Minimal front door (2026-07-10).** Rhino 8 / Grasshopper plugin that exports geometry as
> slicer-ready 3MF packages (Bambu / Orca / Prusa flavors, `ThreeMfWriter`) and can launch the
> target slicer (`SlicerLauncher`). Status: dormant (initial release + icon/sync commits).
> This repo has no AGENTS-DETAIL.md — the key dirs below are the depth (no README).
> **Editing this file: keep it under 7,900 chars** — a size guard alarms via
> `MultiVerse/GOVERNANCE-ALERTS.md` on breach.

## Hard rules
- Do not commit or push without the operator asking.
- Cross-Repo Change Protocol applies — see [`C:\Repos\AGENTS.md`](../AGENTS.md) (umbrella).
- **Component GUIDs are FROZEN locally**: the easy/sequential GUID fixes were already made on the operator's other machine and never pushed — do NOT rewrite, regenerate, or "fix" GUIDs here; it would diverge. Wait for that machine to push (operator constraint, 2026-07).

## Where to look
- `SlicerExport.Core/` — export model + `ThreeMfWriter.cs`, `SlicerLauncher.cs`, `SlicerTarget.cs`.
- `SlicerExport.Grasshopper/` and `SlicerExport.Rhino/` — plugin layers (`SlicerExport.sln`); `SlicerExport.Tests/` — tests.
- `spec/` — design notes; `docs/` is currently untracked.
- Root sample files — `*.3mf` per-slicer cubes, `SampleFile.3dm`, `SampleGHFile.gh` — fixtures for manual checks.
