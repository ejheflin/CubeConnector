# CubeConnector Configuration Wizard — Design

**Date:** 2026-06-24
**Status:** Approved (design); pending implementation plan
**Goal:** Replace the hand-edited `CubeConnectorConfig.json` with an in-Excel wizard that lets non-technical users create, manage, share, and import CubeConnector formulas — backed by silent Power BI auth, dataset enumeration, and model introspection, with no Azure app registration and Pro-license compatibility.

---

## 1. Context & constraints

CubeConnector is a C# **Excel-DNA add-in** (.NET Framework 4.7.2, COM interop). It exposes user-defined functions (UDFs) like `=CC.AmtNet(...)` that query a Power BI semantic model.

Hard constraints that drive this design:

- **UDFs register at load time.** `DynamicFunctionRegistration.AutoOpen()` reads the config and calls `ExcelIntegration.RegisterDelegates(...)` once at Excel startup. New/changed/removed functions therefore **only take effect after an Excel restart**. The wizard must save config and clearly instruct a restart — never fail silently.
- **No Azure AD app registration / no admin consent.** Deployed into locked-down customer tenants. Auth uses well-known first-party client IDs only.
- **Pro-license compatible.** No reliance on the Premium XMLA endpoint.
- **Audience is non-technical.** The UI must be dead simple: business language primary, technical jargon hidden or secondary, sensible defaults, concrete examples.
- **Future team sharing.** Configs must be portable so a colleague in the same tenant can import another user's formulas. The first version ships file-based Import/Export; a "from your team" source layers on later through the same merge path.

### Proven foundation (already built and validated)

- **`PowerBiAuth`** — silent token cascade: WAM zero-click SSO (Windows identity) → DPAPI-cached refresh token → browser auth-code (PKCE, loopback). Account override (`SignInAsDifferentAccount` / `UseWindowsAccount`) persisted via a mode file. No app registration; passes Conditional Access; Pro-compatible. Out-of-process `WamHelper.exe` keeps the native broker out of `EXCEL.EXE`.
- **`PowerBiRestClient`** — enumerates workspaces and datasets via the REST list endpoints (`/myorg/groups`, `/groups/{id}/datasets`, `/datasets`). Pro-compatible. Typed `WorkspaceInfo` / `DatasetInfo`.
- **`ModelIntrospector`** — `IntrospectDataset(datasetId, tenantId)` builds the `pbiazure://` MSOLAP connection from scratch and enumerates tables/fields/measures via DAX `INFO.VIEW.*`. `BuildConnectionString(...)` is the shared connection-string builder.

### Decided trade-off: two separate auth systems

- **Enumeration + introspection (read metadata):** use our silent token.
- **The actual data connection used by the UDFs at query time:** MSOLAP authenticates independently (its own credential cache) — a one-time sign-in per dataset, then cached, exactly like Analyze in Excel. Token injection into MSOLAP was tested and abandoned (the AS layer rejects tokens from borrowed client IDs). This is an accepted, shippable behavior, out of scope for the wizard.

---

## 2. Hosting architecture

**Decision: WebView2 hosted in a WinForms window**, evolving the existing HTML editor (`docs/index.html`) into the product UI.

- A ribbon button (**"Manage Formulas"**) opens a modeless `WizardWindow` (WinForms) hosting a **WebView2** control that renders the HTML/CSS/JS UI.
- A thin **JS ↔ C# bridge** exposes the C# services to the UI. The HTML holds no secrets and no business logic — it is a view only.
- **Dependency:** WebView2 Runtime, which is preinstalled on Windows 11 and ships with Edge/Office (effectively always present for the target audience). Native WinForms was rejected: it would require rebuilding the UI and makes an attractive, plain-language design much harder.

```
Ribbon "Manage Formulas"
  └─▶ WizardWindow (WinForms)
        └─▶ WebView2  ──(bridge)──▶  C# services
                                       ├─ PowerBiAuth        (silent token, account switch)
                                       ├─ PowerBiRestClient  (list workspaces/datasets)
                                       ├─ ModelIntrospector  (measures/fields of a model)
                                       └─ FunctionStore      (functions.json CRUD, import/export, migration)
```

### Bridge API (called from JS, implemented in C#)

| Method | Returns | Notes |
|---|---|---|
| `getAccount()` | `{ upn, mode }` | current identity + wam/browser mode |
| `signInDifferent()` | `{ upn }` | browser sign-in, sets mode=browser |
| `useWindowsAccount()` | `{ upn }` | revert to WAM identity |
| `listDatasets()` | `WorkspaceInfo[]` w/ datasets | from `PowerBiRestClient.GetAllDatasets` |
| `getModel(datasetId, tenantId)` | `{ tables, columns, measures }` | metadata for the builder |
| `getFunctions()` | `UDFConfig[]` | from functions.json |
| `saveFunction(fn)` | ok | upsert by name |
| `deleteFunction(name)` | ok | |
| `exportFunctions(names[])` | file path | writes a shareable JSON |
| `importFunctions(path)` | `{ added, overwritten, skipped }` | merge w/ collision policy |

All bridge calls marshal to/from `functions.json`. Errors return a friendly `{ error }` the UI renders as plain language.

---

## 3. Config store & migration

- **Location:** `%LOCALAPPDATA%\CubeConnector\functions.json` (per-user).
- **Schema:** unchanged `UDFConfig` shape — `functionName`, `tenantId`, `datasetPrefix?`, `datasetId`, `measureName`, `parameters[]` (`name`, `position`, `tableName`, `fieldName`, `dataType`, `filterType`, `isOptional`).
- **`ConfigurationStore.LoadFromJson`** repointed from the `.xll` directory to this per-user path.
- **One-time migration:** on first run, if a legacy `CubeConnectorConfig.json` exists next to the `.xll` and no per-user file exists, copy it to the new location.
- `AutoOpen` continues to read this file and register UDFs at startup (mechanism unchanged).

---

## 4. User flows

### 4a. Consumer flow (hero path — most users)

The least-technical users mainly *receive* formulas a colleague built.

1. Library opens → prominent **"Import formulas someone shared with you."**
2. Pick a `.json` file.
3. Merge with collision handling per function: **overwrite / skip / keep both**.
4. Banner: **"✓ Imported N formulas. Close and reopen Excel to use them."**
5. They never open the builder.

Future: an "Import from your team" source appears beside "from a file," reusing the same merge logic.

### 4b. Builder flow (capable user)

1. Library → **New** (or **Edit** on a card → opens pre-filled).
2. Plain-language editor (section 5).
3. **Save** → banner instructs restart.

### 4c. Library (home)

- Actions: **New**, **Import**, **Export**, account indicator (`signed in: … ▾` with switch-account).
- List of formula cards: friendly name, source model + measure, parameter count, **Edit** / **Delete**.
- Persistent restart banner whenever unsaved-to-Excel changes exist.

---

## 5. The builder — plain-language design

**Labeling rule:** plain phrase **primary**, technical term **muted in parentheses**, **"?" helper** for the actionable explanation. Only three concepts get a "?": *data (model)*, *the number you want (measure)*, *filter*. Helper text says what it is **and how to choose**, with a tiny example. All other technical fields (`dataType`, `position`, `filterType` internals, GUIDs, "UDF") are hidden/auto.

Numbered, top-to-bottom; only ② and ④ are required to save:

- **① What data?** *(model)* ⓘ — dropdown from `listDatasets()`, shown as `Workspace ▸ Model`. GUIDs hidden.
- **② The number you want** *(measure)* ⓘ — dropdown from `getModel()` measures.
- **③ Let people filter by… (optional)** — **"+ Add a filter"** → pick a field → choose **Match value(s)** (= `List`) or **Date range** (auto-creates the From/To pair on one date field = `RangeStart`/`RangeEnd`). `dataType` auto-derived; parameter names auto-suggested and editable; reorder by drag; all optional by default.
- **④ Name it** — user types e.g. `Net Amount`; wizard sanitizes to a valid name and shows `=CC.NetAmount(…)`. The user never learns naming rules.

**Preview block** (never show a bare formula):
- Plain sentence: *"Returns **Net Amount**, filtered by **Account** and a **date range**."*
- Template: `=CC.NetAmount(account, fromDate, toDate)`
- Filled example: `=CC.NetAmount("4000","1/1/2025","12/31/2025")` → *"net amount for account 4000 in 2025."*

**Silent introspection:** the builder reads measures/fields via the Power BI REST **`executeQueries`** endpoint using the silent token already held (running the same `INFO.VIEW` DAX), so selecting a model does **not** trigger a sign-in popup mid-build. *To validate in implementation; `ModelIntrospector` (MSOLAP, one-time prompt) is the proven fallback.*

---

## 6. The restart reality

Because UDFs register only at startup, every add/edit/delete/import shows a persistent, friendly banner:

> **"✓ Saved. Close and reopen Excel to use your changes."**

No state where a newly created formula silently "doesn't work" because Excel wasn't restarted. A **"Restart Excel now"** convenience button is a possible later enhancement, not in v1.

---

## 7. Error handling & edge cases

- **Auth failure** → "Couldn't sign in to Power BI" + Retry; never a raw exception/stack.
- **Model unreadable** (no access / transient) → "Couldn't read this data — you may not have access to it," keep the rest of the UI usable.
- **Save validation** → requires a measure and a name; invalid name characters auto-sanitized; duplicate names prompt overwrite/rename.
- **Import collisions** → explicit per-function choice (overwrite / skip / keep both) or a summary ("3 new, 2 already exist").
- **Export** → user selects which formulas to include in the shared file.
- **Non-semantic-model datasets** (datamarts, dataflow staging, usage-metrics) appear in enumeration; the builder should de-emphasize or gracefully handle ones that can't be introspected.
- **Account/identity mismatch** — WAM uses the Windows logon identity, which may differ from the account that can access a given model; the account switcher resolves this.

---

## 8. Testing strategy

- **C# services** unit-tested independently (already validated via the diagnostic probes that were since removed).
- **Bridge contract** tested against a stub UI (verify each method's shape and error envelope).
- **Manual end-to-end:** build → save → restart → use a formula in a cell; export → import on a second machine/account → restart → use.
- **`executeQueries` validation:** confirm silent introspection returns `INFO.VIEW.*` results on a Pro dataset; fall back to `ModelIntrospector` if not.

---

## 9. Out of scope (v1)

- Central/tenant-hosted config store (file-based Import/Export only; "from team" source later).
- Token injection into the MSOLAP data connection (tested, abandoned).
- "Restart Excel now" automation.
- Editing the underlying DAX measure (the wizard composes filters around an existing measure, it does not author measures).

---

## 10. Components to build

1. **FunctionStore** — per-user `functions.json` CRUD, migration, import/export + merge/collision policy.
2. **ConfigurationStore repoint** — read from per-user path; one-time migration.
3. **WizardWindow + WebView2 host** — WinForms shell, ribbon entry, WebView2 setup.
4. **JS↔C# bridge** — typed methods in section 2, friendly error envelope.
5. **HTML/CSS/JS UI** — evolve `docs/index.html`: library, consumer Import, builder, plain-language labels + "?" helpers, preview block.
6. **executeQueries introspection path** (with `ModelIntrospector` fallback).
7. **Ribbon** — add "Manage Formulas"; remove the temporary smoke-test button.
