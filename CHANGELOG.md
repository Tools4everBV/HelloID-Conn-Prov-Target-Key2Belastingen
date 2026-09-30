# Change Log

All notable changes to this project will be documented in this file. The format is based on [Keep a Changelog](https://keepachangelog.com), and this project adheres to [Semantic Versioning](https://semver.org).

## [2.0.0] - 2026-09-30

### Added

- Marked the Fit For Purpose (FFP) parts, built for the specific requirements of one implementation: the optional delete stored procedure, the optional `MDU_GEBRUIKER` record and the optional menu-group permissions. The README contains an FFP warning and a section listing these parts, and the scripts concerned contain a `FIT FOR PURPOSE (FFP)` note in their header and at the code.
- GitHub workflows to verify the changelog and create a release.
- Full rebuild of the connector to the current HelloID PowerShell V2 script structure (`$actionContext`/`$outputContext`, action messages, audit logs, DryRun support).
- `delete.ps1` (missing from the original connector): soft-delete via the Delete field mapping by default, with an optional customer-specific stored procedure (`DeleteStoredProcedureName`, empty by default; Fit For Purpose (FFP): built for the specific requirements of one implementation, to be reviewed by the consultant per implementation).
- Account import/reconciliation (`import.ps1`) from `WMS_GEBRCODE`, excluding functional/process codes without `GEBR_ORA`.
- Importable `fieldMapping.json` for the managed `WMS_GEBRCODE` account attributes, including the cross-connector Oracle username.
- `uniquenessCheck.ps1` for `GEBRCODE`; the field mapping uses HelloID `Iteration` instead of create-script retry logic.
- Optional `permissions/menuGroups` permissiontype for `MENU_B_GROUP`/`MENU_B_USER`, including permission entitlement import from `MENU_B_USER` (only groups defined in `MENU_B_GROUP`, batches of 500 account references).
- Optional `MDU_GEBRUIKER` creation through the `ManageMduUser` and `MduInDirectorySuffix` settings (disabled by default).
- Governance reconciliation resolutions for Delete and Disable.
- README: Requirements with the Oracle rights per action, Limitations, Not implemented (possible extensions), Common Oracle errors, Script conventions and HelloID Icon URL sections.

### Changed

- README restructured to the current Tools4ever connector layout (supported features, requirements, correlation, field mapping, remarks per topic, database objects).
- Permission scripts moved from the repository root (`permissions.ps1`, `grant_permission.ps1`, `revoke_permission.ps1`) to `permissions/menuGroups/`.
- Retained the `System.Data.OracleClient` provider, because HelloID agent actions run under Windows PowerShell 5.1.
- All actions use the same Oracle connection, query, error handling, splatting and cleanup pattern as the Key2BelastingenOracle connector.
- Enable, disable and delete (without stored procedure) update all action-mapped fields; `INDACTIEF` is controlled by field mapping instead of hardcoded SQL.
- The field mapping creates accounts active (`INDACTIEF = 'J'` on Create), because the `GHG_GEBRCODE` trigger chain rejects an inactive `WMS_GEBRCODE` reference (`##GHG-1005`).
- Delete with a stored procedure only executes the procedure and does not verify its result, so the connector does not depend on the internals of the procedure; the audit log states that the procedure does not return its own logging to HelloID.
- Import filtering (`GEBR_ORA IS NOT NULL`) moved from connector configuration to `import.ps1`.
- Username/password authentication is optional when Oracle integrated security is configured in the Data Source.
- Update uses a normalized `Compare-Object` comparison for all update-mapped mutable fields and reports `NoChanges` when no fields are mapped.
- Update, enable, disable and delete return the current state in `PreviousData` and the expected post-action state in `Data`.

### Fixed

- Oracle errors are detected via the base exception (`GetBaseException()`), so errors wrapped by PowerShell in a `MethodInvocationException` are resolved as Oracle errors.
- Counts use `Measure-Object`, and empty `actionContext.Data`/`outputContext.Data` (no mapped fields, reconciliation) no longer produce a `$null` field name or break the `Compare-Object`/`PreviousData`/`Data` handling.

## [1.0.0] - 2021-06-14

### Added

- Initial work-in-progress version with create, update, enable and disable actions, `menu_b_user` permissions and an `MDU_GEBRUIKER` table insert.
