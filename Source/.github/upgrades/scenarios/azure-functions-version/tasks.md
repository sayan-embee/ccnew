# Tasks — Azure Functions v3 In-Proc → v4 Isolated on .NET 8

## Phase 1 — Foundation (Level 0 projects)

- [ ] **01-common** — Upgrade `CompanyCommunicator.Common` to net8.0; replace deprecated `Microsoft.Azure.Cosmos.Table` → `Azure.Data.Tables`, `Microsoft.Azure.ServiceBus` → `Azure.Messaging.ServiceBus`; update `Azure.Storage.Blobs` (CVE), `Microsoft.Identity.Client`, `Microsoft.Bot.Builder.*`, `Microsoft.Extensions.Configuration`
- [ ] **02-privateblob** — Upgrade `CompanyCommunicator.PrivateBlob` to net8.0; update `Azure.Storage.Blobs` and `Microsoft.Extensions.Configuration`

## Phase 2 — Web Application

- [ ] **03-web** — Upgrade `CompanyCommunicator` (ASP.NET Core) to net8.0; replace `Microsoft.AspNetCore.Authentication.AzureAD.UI` with `Microsoft.Identity.Web`; update auth/SPA packages to 8.0.x; replace `Microsoft.Azure.Storage.Blob` with `Azure.Storage.Blobs`

## Phase 3 — Function Apps (in-proc → isolated)

- [ ] **04-data-func** — Migrate `Data.Func` to .NET 8 isolated worker (smallest, pilot)
- [ ] **05-send-func** — Migrate `Send.Func` to .NET 8 isolated worker (Bot Framework integration)
- [ ] **06-prep-func** — Migrate `Prep.Func` to .NET 8 isolated worker (Durable Functions — largest)

## Phase 4 — Cleanup & Validation

- [ ] **07-validate** — Full solution build, manual smoke checks, document Azure portal deployment settings (`FUNCTIONS_WORKER_RUNTIME=dotnet-isolated`, `netFrameworkVersion=v8.0`)
