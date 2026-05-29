# Migration Plan — Azure Functions v3 In-Proc → v4 Isolated Worker on .NET 8

## 1. Executive Summary

This plan migrates the **Company Communicator** solution (6 projects) from .NET 6 in-process Azure Functions to **.NET 8 LTS isolated worker**. The Functions runtime is already v4; the actual migration work is:

1. Hosting model: **in-process → isolated worker** for the 3 function apps
2. Target framework: **net6.0 → net8.0** for all 6 projects
3. Replace **deprecated/vulnerable NuGet packages** (security-critical: `Azure.Storage.Blobs` CVE)
4. Replace `Microsoft.AspNetCore.Authentication.AzureAD.UI` with `Microsoft.Identity.Web` in the web project

**Why this is required**: The in-process .NET model is end-of-support **November 10, 2026**. The deprecated packages (`Microsoft.Azure.ServiceBus`, `Microsoft.Azure.Cosmos.Table`, `Microsoft.Azure.Storage.Blob`, `Microsoft.Identity.Client` 4.15, `Microsoft.AspNetCore.Authentication.AzureAD.UI`) are no longer maintained.

## 2. Strategy: Incremental Upgrade

**Selected**: Incremental, phased by dependency layer.

**Rationale**: 6 projects, complex dependency graph, **3 function apps with hosting-model migration**, Durable Functions orchestrations (Prep.Func), Bot Framework integration (Send.Func), and deprecated Azure SDK replacements that ripple through the Common library. Risk and blast radius justify phasing.

## 3. Dependency Graph & Phases

```
Level 0 (leaves):
  - CompanyCommunicator.Common              ← used by all 4 above
  - CompanyCommunicator.PrivateBlob         ← used by web app

Level 1 (apps):
  - CompanyCommunicator (AspNetCore)
  - CompanyCommunicator.Send.Func           (Service Bus trigger + Bot Framework)
  - CompanyCommunicator.Prep.Func           (Durable Functions + HTTP + Service Bus)
  - CompanyCommunicator.Data.Func           (Service Bus + Timer triggers)
```

### Phase 1 — Foundation (Level 0)
Upgrade `Common` and `PrivateBlob`: TFM bump, replace deprecated packages, fix code that uses replaced APIs.

### Phase 2 — Web App (Level 1, lower risk)
Upgrade `CompanyCommunicator` (ASP.NET Core): TFM bump, replace `AzureAD.UI` with `Identity.Web`, update auth packages to 8.0.x.

### Phase 3 — Function Apps (Level 1, highest risk)
**One project at a time**, each in this order:
1. `Data.Func` (smallest — 5 functions, no Durable, no Bot Framework)
2. `Send.Func` (medium — 2 functions, Bot Framework)
3. `Prep.Func` (largest — Durable Functions orchestrators + activities + HTTP triggers)

Each function app migration includes:
- TFM net6.0 → net8.0
- `<AzureFunctionsVersion>` stays `v4`
- `<OutputType>Exe</OutputType>` (required for isolated)
- Add `FunctionsSkipCleanOutput` if needed
- Remove `Microsoft.NET.Sdk.Functions`, `Microsoft.Azure.WebJobs.Extensions.*`
- Add `Microsoft.Azure.Functions.Worker`, `Microsoft.Azure.Functions.Worker.Sdk`, and per-trigger extensions
- Replace `Startup.cs` (`FunctionsStartup`) with `Program.cs` (`HostBuilder` + `ConfigureFunctionsWorkerDefaults`)
- Convert each function:
  - `[FunctionName]` → `[Function]`
  - `Microsoft.Azure.WebJobs` → `Microsoft.Azure.Functions.Worker`
  - `ILogger` injected via constructor (or `FunctionContext.GetLogger`)
  - `ExecutionContext` → `FunctionContext`
  - HTTP: `HttpRequest`/`HttpRequestMessage` → `HttpRequestData`/`HttpResponseData`
  - Service Bus: signature change, message body as `string` or `ServiceBusReceivedMessage`
  - Durable Functions: switch to `Microsoft.DurableTask.Client` / `TaskOrchestrationContext`
- Update `host.json` extension bundle version: `[3.*, 4.0.0)` → `[4.*, 5.0.0)`
- Update `local.settings.json`: `FUNCTIONS_WORKER_RUNTIME` = `dotnet-isolated`

### Phase 4 — Cleanup & Validation
Full solution build, smoke tests, verify no leftover deprecated packages, document deployment changes for Azure portal (set `FUNCTIONS_WORKER_RUNTIME=dotnet-isolated`, `linuxFxVersion`/`netFrameworkVersion`).

## 4. Project-by-Project Plans

### 4.1 `Microsoft.Teams.Apps.CompanyCommunicator.Common` (Phase 1)
- TFM: `net6.0` → `net8.0`
- `<LangVersion>8.0</LangVersion>` → remove (use SDK default — C# 12)
- Package updates:
  | Package | Current | Target | Action |
  |---|---|---|---|
  | `Azure.Storage.Blobs` | 12.8.0 | 12.27.0 | **SECURITY FIX** |
  | `Microsoft.Azure.Cosmos.Table` | 1.0.1 | — | **DEPRECATED** → migrate to `Azure.Data.Tables` 12.x |
  | `Microsoft.Azure.ServiceBus` | 4.1.1 | — | **DEPRECATED** → migrate to `Azure.Messaging.ServiceBus` 7.x |
  | `Microsoft.Identity.Client` | 4.15.0 | 4.66+ | **DEPRECATED VERSION** |
  | `Microsoft.Bot.Builder.Integration.AspNet.Core` | 4.12.1 | 4.22+ | Update for .NET 8 |
  | `Microsoft.Extensions.Configuration` | 2.1.1 | 8.0.0 | Align with framework |
  | `Microsoft.Graph` | 3.22.0 | 5.x (or stay on v4 for fewer code changes) | Decide based on impact |
- Code refactor: rewrite `Microsoft.Azure.Cosmos.Table` repository implementations to use `Azure.Data.Tables`; rewrite `Microsoft.Azure.ServiceBus` queue clients to `Azure.Messaging.ServiceBus`.

### 4.2 `Microsoft.Teams.Apps.CompanyCommunicator.PrivateBlob` (Phase 1)
- TFM: `net6.0` → `net8.0`
- Update `Azure.Storage.Blobs` to 12.27.0
- Update `Microsoft.Extensions.Configuration` 3.1.1 → 8.0.0

### 4.3 `Microsoft.Teams.Apps.CompanyCommunicator` (Phase 2)
- TFM: `net6.0` → `net8.0`
- Replace `Microsoft.AspNetCore.Authentication.AzureAD.UI` with `Microsoft.Identity.Web` 4.x → migrate `Startup.cs` auth registration
- Update `Microsoft.AspNetCore.Authentication.JwtBearer` 3.1.4 → 8.0.x
- Update `Microsoft.AspNetCore.Authentication.OpenIdConnect` 3.1.4 → 8.0.x
- Update `Microsoft.AspNetCore.SpaServices.Extensions` 3.1.1 → 8.0.x
- Update `Azure.Storage.Blobs` → 12.27.0 (security)
- Replace `Microsoft.Azure.Storage.Blob` 11.2.3 → use `Azure.Storage.Blobs`
- `Microsoft.ApplicationInsights.AspNetCore` 2.17.0 → 2.22+ (or migrate to `Azure.Monitor.OpenTelemetry`)

### 4.4 `Microsoft.Teams.Apps.CompanyCommunicator.Data.Func` (Phase 3a)
**Smallest function app — pilot for the isolated migration pattern.**
- TFM bump + isolated worker conversion (5 function classes)
- Update `host.json` extension bundle to `[4.*, 5.0.0)`
- Update `local.settings.json` `FUNCTIONS_WORKER_RUNTIME=dotnet-isolated`
- New `Program.cs` replacing `Startup.cs`

### 4.5 `Microsoft.Teams.Apps.CompanyCommunicator.Send.Func` (Phase 3b)
- Apply the pattern from 4.4
- Special: Bot Framework integration — `BotFrameworkHttpAdapter` works in isolated; verify DI registration

### 4.6 `Microsoft.Teams.Apps.CompanyCommunicator.Prep.Func` (Phase 3c)
- Apply the pattern from 4.4
- Special: **Durable Functions** — switch from `Microsoft.Azure.WebJobs.Extensions.DurableTask` to `Microsoft.Azure.Functions.Worker.Extensions.DurableTask`
- Orchestrator & activity attribute changes; `IDurableOrchestrationContext` → `TaskOrchestrationContext`; `IDurableOrchestrationClient` → `DurableTaskClient`

## 5. Package Update Reference (Consolidated)

| Package | Current | Target | Projects Affected | Notes |
|---|---|---|---|---|
| `Azure.Storage.Blobs` | 12.8.0 | 12.27.0 | Common, PrivateBlob, web, Prep.Func, Data.Func | **CVE — security** |
| `Microsoft.AspNetCore.Authentication.AzureAD.UI` | 3.1.1 | — (removed) | web | Replace with `Microsoft.Identity.Web` |
| `Microsoft.AspNetCore.Authentication.JwtBearer` | 3.1.4 | 8.0.10 | web | |
| `Microsoft.AspNetCore.Authentication.OpenIdConnect` | 3.1.4 | 8.0.10 | web | |
| `Microsoft.AspNetCore.SpaServices.Extensions` | 3.1.1 | 8.0.10 | web | |
| `Microsoft.ApplicationInsights.AspNetCore` | 2.17.0 | 2.22.0 | web | |
| `Microsoft.Azure.Cosmos.Table` | 1.0.1 | — | Common | Replace with `Azure.Data.Tables` 12.x |
| `Microsoft.Azure.ServiceBus` | 4.1.1 | — | Common | Replace with `Azure.Messaging.ServiceBus` 7.x |
| `Microsoft.Azure.Storage.Blob` | 11.2.3 | — | web | Replace with `Azure.Storage.Blobs` |
| `Microsoft.Azure.WebJobs.Extensions.DurableTask` | 2.3.0 | — | Prep.Func | Replace with `Microsoft.Azure.Functions.Worker.Extensions.DurableTask` 1.x |
| `Microsoft.Azure.WebJobs.Extensions.ServiceBus` | 3.0.3 | — | Send.Func, Prep.Func, Data.Func | Replace with `Microsoft.Azure.Functions.Worker.Extensions.ServiceBus` 5.x |
| `Microsoft.Azure.WebJobs.Host.Storage` | 3.0.14 | — | Send.Func, Prep.Func, Data.Func | Remove (built-in) |
| `Microsoft.Azure.Functions.Extensions` | 1.1.0 | — | All Func projects | Remove (in-proc DI) |
| `Microsoft.NET.Sdk.Functions` | 3.0.11 | — | All Func projects | Replace with `Microsoft.Azure.Functions.Worker.Sdk` 1.18+ |
| `Microsoft.Extensions.Configuration` | 2.1.1 / 3.1.1 | 8.0.0 | Common, PrivateBlob | |
| `Microsoft.Extensions.Localization` | 3.1.8 | 8.0.10 | Func projects | |
| `Microsoft.Extensions.Localization.Abstractions` | 3.1.8 | 8.0.10 | Func projects | |
| `Microsoft.Extensions.Logging` | 3.1.0 | 8.0.1 | Prep.Func | |
| `Microsoft.Identity.Client` | 4.15.0 | 4.66+ | Common | Used by MSAL flows |
| `Microsoft.VisualStudio.Web.CodeGeneration.Design` | 3.1.5 | 8.0.7 | web | |
| **NEW**: `Microsoft.Azure.Functions.Worker` | — | 1.24.0 | All Func projects | |
| **NEW**: `Microsoft.Azure.Functions.Worker.Sdk` | — | 1.18.1 | All Func projects | |
| **NEW**: `Microsoft.Azure.Functions.Worker.Extensions.ServiceBus` | — | 5.22+ | Send/Prep/Data.Func | |
| **NEW**: `Microsoft.Azure.Functions.Worker.Extensions.Http` | — | 3.3.0 | Prep.Func (HTTP triggers) | |
| **NEW**: `Microsoft.Azure.Functions.Worker.Extensions.Timer` | — | 4.3+ | Data.Func | |
| **NEW**: `Microsoft.Azure.Functions.Worker.Extensions.DurableTask` | — | 1.2+ | Prep.Func | |
| **NEW**: `Microsoft.Azure.Functions.Worker.ApplicationInsights` | — | 2.0+ | All Func projects | |

## 6. Breaking Changes Catalog

### Functions Isolated Worker
- **No more `[FunctionName]`** → use `[Function("Name")]`
- **No more `FunctionsStartup`** → use `HostBuilder` in `Program.cs`
- **No more `ExecutionContext`** → use `FunctionContext`
- **`ILogger` injection** → constructor injection or `FunctionContext.GetLogger<T>()`
- **HTTP triggers** → `HttpRequestData` / `HttpResponseData` instead of `HttpRequest`
- **Service Bus binding signature** changes
- **Durable Functions API** completely renamed (`IDurableOrchestrationContext` → `TaskOrchestrationContext`)

### Azure SDKs
- `Microsoft.Azure.Cosmos.Table.CloudTable` → `Azure.Data.Tables.TableClient`
- `Microsoft.Azure.ServiceBus.QueueClient` / `MessageSender` → `Azure.Messaging.ServiceBus.ServiceBusClient` / `ServiceBusSender`
- `Microsoft.WindowsAzure.Storage.Blob.CloudBlobClient` → `Azure.Storage.Blobs.BlobServiceClient`

### ASP.NET Core 8
- `Microsoft.Identity.Web` registers via `AddMicrosoftIdentityWebApi`/`AddMicrosoftIdentityWebApp` (different signature than `AddAzureADBearer`)

## 7. Testing Strategy

### Per-Project
- ✅ Builds without errors
- ✅ Builds without warnings (StyleCop allowed; new compiler warnings investigated)
- ⚠ Runtime smoke test (manual — no automated tests in this solution)

### Per-Phase
- Solution builds end-to-end after each phase
- Web app loads on `dotnet run`
- Function apps start with `func start --verbose` (or `dotnet run` for isolated)

### Final
- Full solution build clean
- Each function app's startup logs show all triggers registered
- Manual end-to-end test of one notification send (if test environment available)

## 8. Risk Management

| Risk | Likelihood | Impact | Mitigation |
|---|---|---|---|
| Durable Functions API rewrite breaks orchestration semantics | High | High | Migrate `Prep.Func` last; carefully translate each orchestrator/activity; verify with replay scenarios |
| ServiceBus message format incompatibility (peek-lock + dead-letter behavior) | Medium | High | Review each `[ServiceBusTrigger]` signature; test with low-volume queue first |
| `Azure.Data.Tables` query API differences vs `Microsoft.Azure.Cosmos.Table` | Medium | Medium | Repositories in Common are abstracted — focus rewrite there |
| Bot Framework adapter incompatibility with isolated worker | Low | Medium | Bot Framework supports isolated; verify Send.Func independently |
| `Microsoft.Identity.Web` claim mappings differ from `AzureAD.UI` | Medium | Medium | Test login flow carefully; verify role/group claims |
| `Microsoft.Graph` v3 → v5 has major API changes | Low (can stay on v4) | High | **Decision**: Stay on `Microsoft.Graph` 4.x for this migration; defer v5 upgrade |
| No automated test coverage | Certain | High | Manual smoke tests; consider adding integration tests post-migration |

### Rollback
Source is not under git. **Strongly recommend** the user back up the `Source` folder before each phase. We can also keep changes phase-isolated so a single phase can be reverted manually.

## 9. Success Criteria

- [ ] All 6 projects target `net8.0`
- [ ] All 3 function apps run on `dotnet-isolated` worker
- [ ] All deprecated packages removed (Cosmos.Table, ServiceBus 4.x, WindowsAzure.Storage, AzureAD.UI, Identity.Client 4.15)
- [ ] `Azure.Storage.Blobs` security CVE addressed (≥ 12.27.0)
- [ ] Solution builds clean: `dotnet build` returns 0 errors
- [ ] Each function host starts and registers all expected triggers
- [ ] Web app launches and login flow works
- [ ] No remaining `Microsoft.Azure.WebJobs.*` references in function code

## 10. Out of Scope

- Migrating front-end (`ClientApp`) framework versions
- Adding automated test coverage
- Migrating `Microsoft.Graph` to v5 (significant API rewrite — defer)
- Azure infrastructure / ARM / Bicep changes (deployment-side only — documented separately at end)
- Performance tuning
