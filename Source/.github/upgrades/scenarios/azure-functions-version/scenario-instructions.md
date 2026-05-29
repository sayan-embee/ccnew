# Scenario Instructions — Azure Functions Migration (Company Communicator)

## Scope (Confirmed by User)

**Phase A — Compliance-driven migration** (selected over full plan due to risk):

1. ✅ Target framework: `net6.0` → `net8.0` for all 6 projects
2. ✅ Functions hosting model: in-process → **isolated worker** for `Send.Func`, `Prep.Func`, `Data.Func`
3. ✅ Security CVE: `Azure.Storage.Blobs` 12.8.0 → 12.27.0
4. ✅ Replace `Microsoft.AspNetCore.Authentication.AzureAD.UI` → `Microsoft.Identity.Web` (web project)
5. ✅ Update auth/SPA/extension packages to 8.0.x
6. ⏸️ **Deferred to Phase B**: replace `Microsoft.Azure.Cosmos.Table`, `Microsoft.Azure.ServiceBus`, `Microsoft.Azure.Storage.Blob`, `Microsoft.Identity.Client` 4.15. They are deprecated but still functional on .NET 8 — flag as tech debt.

## Preferences

### Flow Mode
**Automatic** — run end-to-end, surface artifacts, only pause when blocked.

### Source Control
- Not a git repository
- User confirmed manual backup is in place

### Backup
- ✅ User confirmed backup of `E:\Company\Company Communicator\.net8\Source` exists.

## Key Decisions Log

- **2025**: User has backup; chose Phase A (compliance-only) over full plan (which would include Cosmos.Table and ServiceBus rewrites). Rationale: no automated tests, no git, hard deadline; full rewrite risk too high.

## Phase B Backlog (Future Work — Not In This Migration)

When you're ready, address these as a separate effort, ideally with test coverage added first:

1. Migrate `Microsoft.Azure.Cosmos.Table` 1.0.1 → `Azure.Data.Tables` 12.x in `Common`
   - Affects all entity classes (TableEntity → ITableEntity)
   - Affects `BaseRepository<T>`, `IRepository<T>`, all 8 derived repositories
   - Affects filter strings (`TableQuery.Generate*` → OData strings)
2. Migrate `Microsoft.Azure.ServiceBus` 4.1.1 → `Azure.Messaging.ServiceBus` 7.x in `Common`
   - Affects `BaseQueue<T>` and all 7 derived queues
   - Coordinate with isolated-worker Service Bus trigger consumers
3. Migrate `Microsoft.WindowsAzure.Storage` 11.2.3 → use `Azure.Storage.Blobs` (web project)
4. Update `Microsoft.Identity.Client` 4.15.0 → 4.66+ (transitive update should help)
5. Consider `Microsoft.Graph` 3.x → 5.x (significant API rewrite — requires planning)
6. Replace deprecated `Microsoft.ApplicationInsights.AspNetCore` with `Azure.Monitor.OpenTelemetry.AspNetCore`
