# CubeConnector Config Wizard — Session Handoff / Status

**Last updated:** 2026-06-24
**Branch:** `feature/config-wizard` (off `master`) — **not merged, not pushed**. 27 commits.
**Repo (real project):** `C:\dev\CubeConnector_gh` (clone of https://github.com/ejheflin/CubeConnector.git).
> Note: `C:\dev\cubeconnector` (lowercase) is a stale, unrelated Office.js scaffold — ignore it. All real work is in `C:\dev\CubeConnector_gh`.

Companion docs: spec `docs/superpowers/specs/2026-06-24-cubeconnector-wizard-design.md`; plan + task state `docs/superpowers/plans/2026-06-24-cubeconnector-wizard.md(.tasks.json)`.

---

## Goal (achieved)
Replace the hand-edited `CubeConnectorConfig.json` with an **in-Excel wizard** that lets non-technical users create / edit / import / share Power BI "formulas" (UDFs) — backed by **silent auth, dataset enumeration, and model introspection, with NO Azure AD app registration, NO admin consent, and Pro-license compatibility.**

CubeConnector is a **C# Excel-DNA add-in** (.NET Framework 4.7.2, COM interop), classic `packages.config`.

---

## Build / run / verify

**Build** (close Excel first — it locks the `.xll`; watch for orphan `EXCEL.EXE`):
```
"C:\Program Files (x86)\Microsoft Visual Studio\18\BuildTools\MSBuild\Current\Bin\MSBuild.exe" \
  "C:\dev\CubeConnector_gh\CubeConnector\CubeConnector.csproj" /t:Rebuild /p:Configuration=Debug /p:Platform=AnyCPU /v:minimal /nologo
```
(There is one pre-existing harmless warning `CS0168` in `CacheManager.cs`. No test harness — verification is build-green + manual in Excel.)

**WamHelper** (separate SDK-style net472 exe, built with `dotnet`, NOT the BuildTools MSBuild which lacks the SDK resolver):
```
dotnet build "C:\dev\CubeConnector_gh\WamHelper\WamHelper.csproj" -c Release
```
The add-in's post-build target copies `WamHelper\bin\Release\*` to `bin\Debug\WamHelper\` and `ui\*` + `WebView2Loader.dll` beside the add-in.

**Load in Excel:** add the **non-packed** `…\CubeConnector\bin\Debug\CubeConnector-AddIn64.xll` (so WebView2 DLLs, `WebView2Loader.dll`, and `ui\` resolve beside it). Ribbon: **Data ▸ CubeConnector ▸ Manage Formulas** opens the docked task pane.

**Per-user config:** `%LOCALAPPDATA%\CubeConnector\functions.json` (+ `pbi_refresh.bin` DPAPI token cache, `pbi_mode.txt` account mode, `WebView2\` profile).

---

## Architecture (key files in `CubeConnector/`)
- **PowerBiAuth.cs** — token cascade: WAM zero-click (out-of-proc `WamHelper.exe`) → DPAPI-cached refresh → browser auth-code+PKCE (loopback). `AccountMode` wam/browser persisted; `SignInAsDifferentAccount` / `UseWindowsAccount`. ⚠ Has a hardcoded `WamHelperDevPath` fallback to remove for production.
- **WamHelper/** — SDK-style net472 console exe; MSAL + broker; `silent`/`interactive` modes; prints `ACCESS_TOKEN=` / `ERROR=`.
- **PowerBiRestClient.cs** — REST enumerate (`/myorg/groups`, `/datasets`) + `ExecuteQueriesIntrospect` (runs `INFO.VIEW.TABLES/COLUMNS/MEASURES` via `executeQueries`, incl. Description). Dataset cache (`WarmDatasetCache`/`GetAllDatasetsCached`/`ClearCache`).
- **ModelIntrospector.cs** — `BuildConnectionString(datasetId, tenantId)` (uses the FIXED AiE client id — see below), `IntrospectDataset`, `Introspect` (rides existing connection). `AnalyzeInExcelClientId` const.
- **ModelMetadata.cs** — DTO (Tables / Columns{Table,Name,DataType,IsHidden,Description} / Measures{Table,Name,Description}).
- **FunctionStore.cs** — per-user `functions.json`: GetAll/Save/Delete, SanitizeName, Export, Import(+merge), MigrateLegacyIfNeeded.
- **ConfigurationStore.cs** — `GetAllConfigs()` delegates to FunctionStore; `Invalidate()`.
- **DynamicFunctionRegistration.cs** — `AutoOpen` registers UDFs; `ReloadFunctions()` registers NEW funcs at runtime via `ExcelAsyncUtil.QueueAsMacro` (restart only for deletes/arity changes); `EnsureConnectionForDataset` / `ConnectionNameForDataset(config)` / `ShortDatasetId` (connection-per-dataset).
- **RefreshManager.cs** — refresh partitions cells by `Config.DatasetId`, runs the pooling pipeline per dataset against that dataset's connection + hidden query sheet (`ProcessCellGroup`, `ExecuteBatchQuery(query, connName, sheet, listObj)`).
- **DrillthroughManager.cs / PivotManager.cs** — resolve the per-dataset connection for the cell's function.
- **WizardBridge.cs** — COM-visible JSON facade: GetAccount/SignInDifferent/UseWindowsAccount/ListDatasets/GetModel/GetFunctions/SaveFunction/DeleteFunction/ReloadFunctions/Export/Import. Per-dataset `_modelCache`.
- **WizardWindow.cs** (popup fallback) / **WizardPaneControl.cs** (docked CTP, COM-visible w/ GUID) — host WebView2; `PreferredColorScheme = OfficeTheme.Scheme()`.
- **OfficeTheme.cs** — maps `HKCU\Software\Microsoft\Office\16.0\Common\UI Theme` (Black→Dark, White/Colorful/DarkGray→Light, else Auto) so the pane follows the **Office** theme, not the OS.
- **CubeConnectorRibbon.cs** — `Manage Formulas` button toggles the CTP (falls back to popup window); fires dataset prefetch on click.
- **ui/** (`index.html`, `app.js`, `styles.css`) — WebView2 UI: library + builder; searchable collapsible pickers (model grouped by workspace, measure/field grouped by table, descriptions as subheadings); Excel **formula-bar** signature; light/dark; logo cyan→violet accent used sparingly; drag-to-reorder filters; instant measure pre-fill; calls `cc` host object.

---

## Critical findings / decisions (don't re-litigate)
1. **No XMLA / Pro-compatible.** Use `pbiazure://api.powerbi.com` MSOLAP (Analyze-in-Excel transport), NOT the Premium `powerbi://` XMLA endpoint.
2. **No app registration.** Auth uses the well-known **Azure CLI public client** `04b07795-8ddb-461a-bbee-02f9e1bf7b46`. **Device-code flow is blocked by this tenant's Conditional Access; browser auth-code passes.** WAM (broker) gives true zero-click and also works with the borrowed client id.
3. **The MSOLAP `Identity Provider` third GUID is the FIXED Analyze-in-Excel client id `929d0ec0-7a41-4b1e-bc7c-b754a28bddcc` — NOT the tenant.** Putting a tenant there → `AADSTS700016`. `tenantId` in config is now vestigial.
4. **Two separate auth systems:** our token (REST enumeration/introspection) vs. MSOLAP's own ClaimsToken auth for the data connection (one-time prompt per dataset, then cached). **Token injection into MSOLAP was tested and abandoned** (AS layer rejects borrowed-appid tokens).
5. **Live UDF reload:** Excel-DNA can register NEW functions at runtime only via `QueueAsMacro` (macro context); it cannot unregister or re-arity — deletes/param-count changes need a restart.
6. **Connection-per-dataset:** required because the wizard allows multiple models; the old single shared connection only served `configs[0]`.

---

## Verified live (by the user)
Silent auth (WAM + browser fallback); enumerate datasets; silent introspection (no popup); build→save→**formula resolves in a cell with no restart**; **multi-model** (two datasets resolve correctly); readable connection names; modern UI in **light + dark** following the Office theme; searchable/collapsible pickers w/ descriptions; drag-reorder; instant pre-fill.

## Remaining productionization (none block daily use)
1. **Import/Export use `prompt()` for file paths** — replace with native `OpenFileDialog`/`SaveFileDialog` (add bridge methods).
2. **Remove `PowerBiAuth.WamHelperDevPath`** dev fallback; ensure installer ships `WamHelper\` beside the `.xll`.
3. **Exercise import/export round-trip + account-switch** end-to-end (built, lightly tested).
4. **Merge/PR** `feature/config-wizard` → `master` (not done; user merges on request). Consider the `finishing-a-development-branch` flow.
5. Optional: model-name back-fill for pre-existing functions (their connections show `CubeConnector Data (<id>)` until edited+saved once).

## How to resume
- Read this file + the spec/plan in `docs/superpowers/`.
- Task state: `docs/superpowers/plans/2026-06-24-cubeconnector-wizard.md.tasks.json` (Tasks 1–13 complete).
- Everything is committed on `feature/config-wizard`; `git log master..feature/config-wizard` shows the 27 commits.
