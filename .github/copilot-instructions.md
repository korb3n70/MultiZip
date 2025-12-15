Purpose
-------
This file gives focused, actionable guidance for AI coding agents working on the MultiZip repository so they can be productive immediately.

Quick Summary
-------------
- Small, single-binary C# utility with a PowerShell wrapper. Key files: `MultiZIP.CS`, `MultiZIP2.cs`, `MultiZip.ps1`, `README.md`.
- Prioritize minimal, targeted changes: this repo contains a few source files rather than a multi-project solution.

Big picture / architecture
--------------------------
- Single-process CLI utility implemented in C# (see `MultiZIP.CS` and `MultiZIP2.cs`). Expect a `Main` entry and direct file-system operations.
- `MultiZip.ps1` is a convenience wrapper/workflow for running and/or packaging the binary on Windows — use it to understand expected invocation patterns and runtime environment.

Key files
---------
- `MultiZIP.CS` / `MultiZIP2.cs`: Main implementation. Look here for argument parsing, file enumeration, compression logic, and exit codes.
- `MultiZip.ps1`: PowerShell wrapper. Inspect for example invocations, environment assumptions (paths, elevated permissions), and any build/run steps.
- `README.md`: Project intent and any manual usage examples. Use it for quick context if present.

Developer workflows (what to run)
--------------------------------
- If there is no `.csproj`, compile single-file C# sources on Windows using the platform toolchain, for example:

  - Use the PowerShell wrapper when present: open `MultiZip.ps1` to see how the repo author runs the tool — it often encodes required arguments or environment setup.

  - Fallback compile (when no project file exists): `csc.exe /t:exe /out:MultiZip.exe MultiZIP.CS` (or `MultiZIP2.cs`) or use Visual Studio / `dotnet` tooling if a `.csproj` is added.

Patterns & conventions (specific to this repo)
---------------------------------------------
- Small-file, imperative C# style: avoid adding heavy abstractions unless adding significant features.
- CLI-first behavior: prefer explicit, synchronous filesystem operations over background services.
- Backward-compatibility with existing command-line invocation: when modifying argument parsing, preserve flags used by `MultiZip.ps1` and examples in `README.md`.

Integration points & assumptions
--------------------------------
- Windows-first expectations: presence of a `.ps1` runner indicates Windows usage and PowerShell-based workflows.
- No obvious external network or database integrations; focus on local filesystem and OS APIs.

What to change and how (actionable rules for AI editing sessions)
----------------------------------------------------------------
- When adding features, update or add an example invocation in `MultiZip.ps1` and extend `README.md` with the new usage.
- Preserve command-line behavior: keep exit codes numeric and consistent; callers (scripts) may rely on them.
- Keep changes focused to the few C# files; do not scaffold a large solution unless requested.

Examples / where to look
------------------------
- To understand runtime arguments: open `MultiZIP2.cs` and search for `Main(` or argument-parsing logic.
- To validate typical invocation: open `MultiZip.ps1` to see how the binary is called and any environment variables used.

When you can't find something
----------------------------
- If a `.csproj` or build file is missing, rely on `csc` or the `MultiZip.ps1` wrapper for invocation guidance. Ask the repo owner whether they prefer `dotnet` SDK projects.

Questions for the maintainer (when unsure)
-----------------------------------------
- Should we create a `.csproj` and adopt `dotnet build` as the canonical build flow?
- Are both `MultiZIP.CS` and `MultiZIP2.cs` active implementations or is one legacy? Prefer the one referenced by `MultiZip.ps1` and `README.md`.

Feedback
--------
If any section is unclear or you want more concrete examples (e.g., exact compile commands or argument names), tell me which file to inspect and I will extract exact snippets.
