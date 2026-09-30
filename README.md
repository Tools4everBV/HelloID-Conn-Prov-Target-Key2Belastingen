# HelloID-Conn-Prov-Target-Key2Belastingen

> [!IMPORTANT]
> This repository contains the connector and configuration code only. The implementer is responsible to acquire the connection details such as the Oracle data source, a service account and the Oracle Client. You might need to sign a contract or agreement with the supplier (Centric) before implementing this connector. Please contact the client's application manager and database administrator to coordinate the connector requirements.

<p align="center">
  <img src="https://raw.githubusercontent.com/Tools4everBV/HelloID-Conn-Prov-Target-Key2Belastingen/refs/heads/main/Logo.png" width="500">
</p>

> [!WARNING]
> **Fit For Purpose (FFP)**
>
> This connector contains **Fit For Purpose (FFP)** parts: parts that were built for the specific requirements of one implementation, of which it is not certain that they are generic or fit the next implementation. Examples are the optional delete stored procedure and the optional menu-group permissions. These parts are marked with `FIT FOR PURPOSE (FFP)` in the scripts concerned and listed in [Fit For Purpose (FFP)](#fit-for-purpose-ffp). Review them against the requirements of the customer before implementing this connector.

> [!WARNING]
> This connector writes **directly to the Key2 Belastingen tables** in the Oracle database; there is no supplier API. Keep the following in mind:
>
> - A person must already have a native Oracle database user (managed by the [Key2BelastingenOracle](https://github.com/Tools4everBV/HelloID-Conn-Prov-Target-Key2BelastingenOracle) connector), see [Introduction](#introduction)
> - Accounts are always created **active**, because of a trigger in Key2 Belastingen, see [Create behavior](#create-behavior)
> - Delete is a soft-delete by default; the optional delete stored procedure is customer-specific and must be reviewed per implementation, see [Delete behavior](#delete-behavior)
> - Functional Key2 Belastingen rights (other than the optional menu groups) are not managed, see [Function-specific rights](#function-specific-rights)
>
> Read the [Supported features](#supported-features) table and the [Remarks](#remarks) section carefully and test on a non-production database first.

## Table of contents

- [HelloID-Conn-Prov-Target-Key2Belastingen](#helloid-conn-prov-target-key2belastingen)
  - [Table of contents](#table-of-contents)
  - [Introduction](#introduction)
  - [Supported features](#supported-features)
  - [Getting started](#getting-started)
    - [HelloID Icon URL](#helloid-icon-url)
    - [Requirements](#requirements)
    - [Connection settings](#connection-settings)
    - [Correlation configuration](#correlation-configuration)
    - [Field mapping](#field-mapping)
    - [Account reference](#account-reference)
    - [Uniqueness](#uniqueness)
    - [Permissions](#permissions)
  - [Remarks](#remarks)
    - [Fit For Purpose (FFP)](#fit-for-purpose-ffp)
    - [Limitations](#limitations)
    - [Not implemented (possible extensions)](#not-implemented-possible-extensions)
    - [Combination with the Key2BelastingenOracle connector](#combination-with-the-key2belastingenoracle-connector)
    - [Create behavior](#create-behavior)
    - [Update behavior](#update-behavior)
    - [Enable and disable behavior](#enable-and-disable-behavior)
    - [Delete behavior](#delete-behavior)
    - [Menu-group permissions](#menu-group-permissions)
    - [Function-specific rights](#function-specific-rights)
    - [MDU integration](#mdu-integration)
    - [Account import](#account-import)
    - [Governance reconciliation resolutions](#governance-reconciliation-resolutions)
    - [Integrated security](#integrated-security)
    - [Query execution](#query-execution)
    - [Common Oracle errors](#common-oracle-errors)
  - [Development resources](#development-resources)
    - [Database objects](#database-objects)
    - [Script conventions](#script-conventions)
  - [Getting help](#getting-help)
  - [HelloID docs](#helloid-docs)

## Introduction

_HelloID-Conn-Prov-Target-Key2Belastingen_ is a _target_ connector. Key2 Belastingen (Centric) is a municipal tax application that runs on an Oracle database. This connector manages the **Key2 Belastingen application account**: a record in the `WMS_GEBRCODE` table.

Every `WMS_GEBRCODE` record references a native Oracle database user in `GEBR_ORA`. That Oracle user must exist before the record can be created, either via the separate [HelloID-Conn-Prov-Target-Key2BelastingenOracle](https://github.com/Tools4everBV/HelloID-Conn-Prov-Target-Key2BelastingenOracle) connector (recommended) or via another (e.g. manual) process. Recommended provisioning order: Oracle user first, then this connector. See [Combination with the Key2BelastingenOracle connector](#combination-with-the-key2belastingenoracle-connector).

## Supported features

The following features are available:

✅ = implemented, ✅⚠️ = implemented with limitations, ❌ = not implemented. See [Limitations](#limitations) and [Not implemented (possible extensions)](#not-implemented-possible-extensions).

| Feature | Supported | Actions | Remarks |
| ------- | --------- | ------- | ------- |
| Account Lifecycle | ✅⚠️ | Create, Update, Enable, Disable, Delete | Create is always active. Update only changes `GEBR_OMS` and `EMAIL` (when mapped for Update). Delete is a soft-delete via field mapping (`INDACTIEF = 'N'`), or optionally a customer-specific stored procedure. No record is ever removed. [Read more](#delete-behavior) |
| Uniqueness | ✅ | - | `GEBRCODE`, generated by field mapping with `Iteration` and validated by `uniquenessCheck.ps1`. [Read more](#uniqueness) |
| Permissions | ✅⚠️ | Retrieve, Grant, Revoke | Optional menu-group permissions via `permissions/menuGroups` (FFP); install only where menu groups can be governed by HelloID Business Rules. Functional rights are not managed. [Read more](#menu-group-permissions) |
| Resources | ❌ | - | Not applicable. [Read more](#not-implemented-possible-extensions) |
| Entitlement Import: Accounts | ✅ | - | From `WMS_GEBRCODE`, excluding functional/process codes without `GEBR_ORA`. [Read more](#account-import) |
| Entitlement Import: Permissions | ✅⚠️ | - | Menu-group assignments from `MENU_B_USER` (only when the optional permissiontype is installed). [Read more](#menu-group-permissions) |
| Governance Reconciliation Resolutions | ✅ | Delete, Disable | Hardcoded values in reconciliation mode. [Read more](#governance-reconciliation-resolutions) |

## Getting started

### HelloID Icon URL

URL of the icon used for the HelloID Provisioning target system.

```
https://raw.githubusercontent.com/Tools4everBV/HelloID-Conn-Prov-Target-Key2Belastingen/refs/heads/main/Icon.png
```

### Requirements

Before implementing this connector, ensure the following requirements are met.

**HelloID agent:**

- **Local agent**:<br>
  The connector only works through a local HelloID agent; the cloud agent cannot reach the database or use the Oracle Client.
- The agent needs network access to the Oracle database server, the same database as the Key2BelastingenOracle connector (default port `1521`, or `2484` for TCPS)
- Windows PowerShell 5.1 and a compatible Oracle Client installed on the HelloID agent server. The connector uses the .NET Framework `System.Data.OracleClient` provider, which is part of Windows PowerShell 5.1 (not PowerShell 7) and is deprecated by Microsoft, but is the only option that works in the HelloID agent actions
- The architecture of the Oracle Client (32-bit or 64-bit) must match the architecture of the PowerShell process of the agent (normally 64-bit)

**Oracle service account:**

Use a dedicated Oracle service account (e.g. `SVC_HelloID`) and grant only the rights below. The service account needs **table rights on the Key2 Belastingen tables only**; it needs no DBA rights and no rights to manage Oracle users (that is done by the Key2BelastingenOracle connector). The table lists per right which connector actions need it, so rights can be left out when an action is not used.

| Right | Required for | Remarks |
| ----- | ------------ | ------- |
| `CREATE SESSION` | All actions | The service account must be able to log in |
| `SELECT` on `WMS_GEBRCODE` | All account actions, correlation, uniqueness check, account import, permissions | Look up the record |
| `INSERT` on `WMS_GEBRCODE` | Create | |
| `UPDATE` on `WMS_GEBRCODE` | Update, enable, disable, delete (soft-delete) | Not needed for delete when `DeleteStoredProcedureName` is used |
| `INSERT` on `MDU_GEBRUIKER` | Create, only when `ManageMduUser` is enabled ([FFP](#fit-for-purpose-ffp)) | |
| `EXECUTE` on the delete stored procedure | Delete, only when `DeleteStoredProcedureName` is configured ([FFP](#fit-for-purpose-ffp)) | Plus the rights the procedure itself needs; check this with the database administrator |
| `SELECT` on `MENU_B_GROUP` | Permissions list and import, only for the optional `menuGroups` permissiontype ([FFP](#fit-for-purpose-ffp)) | |
| `SELECT`, `INSERT` on `MENU_B_USER` | Grant menu group, permission import | |
| `SELECT`, `DELETE` on `MENU_B_USER` | Revoke menu group | |

Example of the grants for a service account that only manages the `WMS_GEBRCODE` records (no optional parts). Adjust to the actions and optional parts in scope:

```sql
GRANT CREATE SESSION TO SVC_HELLOID;
GRANT SELECT, INSERT, UPDATE ON WMS_GEBRCODE TO SVC_HELLOID;
```

> [!WARNING]
> - The insert and update triggers of Key2 Belastingen (for example the trigger chain that creates the `GHG_GEBRCODE` record) run as part of the statement. Test in a non-production environment that the service account can create, update and disable a record with exactly the rights above
> - The service account needs no rights to create or drop Oracle users, and cannot grant Oracle roles or privileges through this connector
> - Rights on tables of another schema must be granted on the owner's objects (for example `GRANT ... ON <owner>.WMS_GEBRCODE`); when the tables are not in the schema of the service account, confirm how the table names resolve (synonyms or schema prefix). The scripts use the unqualified table names

**Oracle database and Key2 Belastingen:**

- The Oracle user of each person must exist before the record is created; this is normally the Oracle user created by the Key2BelastingenOracle connector
- `GEBRCODE` can be at most 6 characters, see [Uniqueness](#uniqueness)
- The optional fields of `WMS_GEBRCODE` (`LOCATIE`, `VRIJ_VELD`, `SUBJECTNR`, the `IND_*` indicators, `EXE_USER`, `HASHEE`) are customer specific; confirm their usage with the application manager

**Other connectors:**

- A Key2BelastingenOracle target system (or another source of the Oracle username) for the `GEBR_ORA` and `GEBRCODE` mappings and for correlation

### Connection settings

The following settings are required to connect to the Oracle database.

| Setting | Description | Mandatory |
| ------- | ----------- | --------- |
| DataSource | Oracle connection string data source, format `[server]:[port]/[service]`. Append `;Integrated Security=yes` to use integrated security. | Yes |
| Username | Oracle service account username. Leave Username and Password empty for integrated security. | No |
| Password | Oracle service account password. Leave Username and Password empty for integrated security. | No |
| DeleteStoredProcedureName | **Optional, empty by default.** Customer-specific stored procedure executed on delete instead of the soft-delete via field mapping. Read [Delete behavior](#delete-behavior) before using it. | No |
| ManageMduUser | **Optional, disabled by default.** Create a legacy `MDU_GEBRUIKER` record after the `WMS_GEBRCODE` record. [Read more](#mdu-integration) | No |
| MduInDirectorySuffix | Suffix appended to the lowercase Oracle username for `MDU_IN_DIR` when `ManageMduUser` is enabled (default `\ACC\InFiles`). | No |

### Correlation configuration

The correlation configuration is used to specify which properties are used to match an existing account within Key2 Belastingen to a person in HelloID.

| Setting | Value |
| ------- | ----- |
| Enable correlation | `True` |
| Person correlation field | The Oracle username of the person, e.g. `Accounts.Key2BelastingenOracle.USERNAME` (use the system name of the installed Key2BelastingenOracle target system) |
| Account correlation field | `GEBR_ORA` |

> [!NOTE]
> Create fails with a clear message when the correlation value is empty. This usually means the person has no Oracle user (yet).

### Field mapping

The field mapping can be imported by using the `fieldMapping.json` file. The supplied mapping contains:

| Field | Actions | Value |
| ----- | ------- | ----- |
| `GEBRCODE` | Create | First 6 characters of the Oracle username, with `Iteration` as suffix on conflicts. [Read more](#uniqueness) |
| `GEBR_ORA` | Create | Oracle username of the Key2BelastingenOracle account |
| `GEBR_OMS` | Create, Update | Display name (max. 40 characters) |
| `EMAIL` | Create, Update | `Person.Contact.Business.Email` |
| `INDACTIEF` | Create, Enable: `J`; Disable, Delete: `N` | Active indicator. [Read more](#create-behavior) |
| `IND_JURIST`, `IND_JB`, `IND_DW` | Create | Fixed `N` |
| `SUBJECTNR`, `TELNR`, `FAXNR`, `LOCATIE`, `VRIJ_VELD`, `EXE_USER`, `HASHEE` | Create | Fixed empty |

> [!NOTE]
> Replace `Key2BelastingenOracle` in `Person.Accounts.Key2BelastingenOracle` in the `GEBR_ORA` and `GEBRCODE` mappings with the system name of the installed Key2BelastingenOracle target system. Confirm the usage of the optional fields per implementation.

### Account reference

The account reference is the `GEBRCODE` (the primary key of `WMS_GEBRCODE`).

### Uniqueness

`GEBRCODE` is generated in field mapping. When `uniquenessCheck.ps1` reports it as non-unique in `WMS_GEBRCODE`, HelloID increments `Iteration` and reruns the mapping. The create action does not generate or retry account codes itself. Configure `GEBRCODE` as a unique field on the target system. Create fails when `GEBRCODE` is empty or longer than 6 characters.

### Permissions

The connector includes the optional `permissions/menuGroups` permissiontype, based on `MENU_B_GROUP` and `MENU_B_USER`. See [Menu-group permissions](#menu-group-permissions).

## Remarks

### Fit For Purpose (FFP)

Fit For Purpose (FFP) means that a part was built for the specific requirements and wishes of one customer/implementation, and is not necessarily suitable for the next customer/implementation. The scripts concerned contain a `FIT FOR PURPOSE (FFP)` note in their header and at the code itself.

The following parts are FFP and must be reviewed per implementation:

| Part | Script / file | Review |
| ---- | ------------- | ------ |
| Delete stored procedure (`DeleteStoredProcedureName`) | `delete.ps1`, `configuration.json` | Calling a customer-specific stored procedure on delete was built for one implementation. Not every implementation has such a procedure, and what it does differs per customer. Without a procedure, delete performs a generic soft-delete. See [Delete behavior](#delete-behavior). |
| `MDU_GEBRUIKER` record (`ManageMduUser`) | `create.ps1`, `configuration.json` | Comes from a previous implementation and is disabled by default. It is not certain that every implementation uses `MDU_GEBRUIKER` or this record layout. See [MDU integration](#mdu-integration). |
| Menu-group permissions | `permissions/menuGroups/*.ps1` | Come from a previous implementation and were not used in the implementation this connector was rebuilt for. It is not certain that every implementation manages menu groups this way. Test before use. See [Menu-group permissions](#menu-group-permissions). |
| Create active (`INDACTIEF = 'J'`) | `fieldMapping.json` | Based on the triggers observed in one Key2 Belastingen installation. See [Create behavior](#create-behavior). |
| Optional `WMS_GEBRCODE` fields | `fieldMapping.json` | The usage of fields such as `LOCATIE`, `VRIJ_VELD`, `SUBJECTNR` and the `IND_*` indicators differs per customer. Confirm with the application manager. |

### Limitations

Read this section before implementing. These are the technical (im)possibilities of this connector:

| Limitation | Explanation |
| ---------- | ----------- |
| Only the application account | Only `WMS_GEBRCODE` (and optionally `MDU_GEBRUIKER` and `MENU_B_USER`) is managed. The Oracle user itself is managed by the Key2BelastingenOracle connector; this connector needs that username (`GEBR_ORA`) and does not create it |
| No inactive create | Accounts are always created active, see [Create behavior](#create-behavior) |
| No real delete | Delete is a soft-delete (`INDACTIEF = 'N'`) or a customer-specific stored procedure (FFP). `WMS_GEBRCODE` records are never removed and a created `MDU_GEBRUIKER` record is not removed |
| Only mapped fields change | Enable, disable and delete only set the fields mapped for those actions (by default `INDACTIEF`). Start and end dates are not evaluated by the application itself; HelloID triggers enable and disable |
| Update is limited | Update only changes the fields mapped for Update (by default `GEBR_OMS` and `EMAIL`). `GEBRCODE` and `GEBR_ORA` are set at create and are not updated; a changed Oracle username does not lead to an update |
| Function-specific rights are not managed | `MWB_GEBR_GROEP`, `MWB_STUUR`, `WMS_GEBRWVD` and `GHG_GEBRCODE` are not managed, see [Function-specific rights](#function-specific-rights) |
| Permissions are limited to menu groups | Optional and FFP; there are no other permissiontypes |
| Direct table access | There is no supplier API; the connector writes to the tables and the triggers of Key2 Belastingen run. The behavior of those triggers can differ per installation and version |
| Delete order | When a stored procedure is used, run this delete **before** the Key2BelastingenOracle delete |
| No rollback | Statements are committed directly by Oracle; a failed action is not rolled back |
| Agent requirements | A local HelloID agent, Windows PowerShell 5.1 and an Oracle Client are required; PowerShell 7 and the cloud agent are not supported |
| Reconciliation values are hardcoded | See [Governance reconciliation resolutions](#governance-reconciliation-resolutions) |

### Not implemented (possible extensions)

The following is possible but was not built. A next implementation can add it when needed.

| Not implemented | How | Reason |
| --------------- | --- | ------ |
| Functional rights (`MWB_GEBR_GROEP`, `MWB_STUUR`, `WMS_GEBRWVD`, `GHG_GEBRCODE`) | Extra permission scripts or create logic | Requires customer-specific mapping or copy logic, see [Function-specific rights](#function-specific-rights) |
| Inactive create | - | Not possible in the observed installation because of the trigger chain, see [Create behavior](#create-behavior) |
| Removal of the record (hard delete) | `DELETE` in `delete.ps1` | Not built: dependent records (for example `GHG_GEBRCODE`) and audit requirements are customer specific; use the delete stored procedure (FFP) when needed |
| Removal of the `MDU_GEBRUIKER` record on delete | `DELETE` in `delete.ps1` | The record is optional (FFP) and its usage differs per customer |
| Update of other attributes (for example `LOCATIE`, `VRIJ_VELD`) | Map them for Update in `fieldMapping.json` | Only `GEBR_OMS` and `EMAIL` were needed. Mapped mutable fields are compared and updated without script changes |

> [!NOTE]
> The reasons above are derived from the current implementation. Confirm them per implementation with the customer and the application manager.

### Combination with the Key2BelastingenOracle connector

- Provisioning order: Active Directory (username), then the Key2BelastingenOracle connector (Oracle user), then this connector (`WMS_GEBRCODE` record referencing the Oracle username)
- Delete order: when a delete stored procedure is used, run this delete **before** the Key2BelastingenOracle delete, as such a procedure may require the Oracle user to still exist
- Configure this connector as a dependent system of the Key2BelastingenOracle target system, so the execution order is guaranteed and the Oracle username is available. See [Dependent systems](https://docs.helloid.com/en/provisioning/target-systems/share-account-fields-between-target-systems/access-shared-target-account-fields.html)
- The Oracle user needs a role or system privilege (for example `CREATE SESSION`) assigned by the Key2BelastingenOracle connector to be able to log in; this connector does not assign that

### Create behavior

`INDACTIEF` values: `J` = active, `N` = inactive. Accounts are always created active (`INDACTIEF = 'J'`).

> [!WARNING]
> An inactive create is not possible. In the observed Key2 Belastingen installation, the insert trigger `WMS_WGBR_AIS_GHG` automatically creates a `GHG_GEBRCODE` record, whose trigger `GHG_GGEB_BIR` requires an active `WMS_GEBRCODE` reference. An inactive create fails with `ORA-20000: ##GHG-1005: Verwijzing GHG_F_GGEB_WGBR bestaat niet of is niet actief`. The account is therefore active from creation, not from the start date or Enable.

- Create uses all fields mapped for Create as columns of the `INSERT`
- When an account with the same `GEBR_ORA` already exists, it is correlated instead of created
- `GEBRCODE` is uppercased, must not be empty and may be at most 6 characters
- Create fails when the correlation value (the Oracle username) is empty

### Update behavior

- Update compares all update-mapped, mutable fields with `Compare-Object` and only updates changed values. Null and empty strings are normalized to the same value before comparison
- When no fields are mapped for Update, update reports `NoChanges`
- `GEBRCODE` and `GEBR_ORA` are not updated
- Update, enable, disable and delete populate `outputContext.PreviousData` with the current database state and `outputContext.Data` with the expected post-action state
- Update, enable, disable and delete fail when the record no longer exists (delete is skipped instead)

### Enable and disable behavior

- Enable and disable update every field mapped for their respective action (the supplied mapping only sets `INDACTIEF`: `J` and `N`). They fail with a clear message when no fields are mapped. Additional action-specific fields can be added in field mapping without changing the scripts
- The update is skipped (`NoChanges`) when the values already match
- Enable or disable of the record does not lock or unlock the Oracle user; that is done by the Key2BelastingenOracle connector
- Existing sessions are not ended

### Delete behavior

By default (`DeleteStoredProcedureName` empty), delete performs a **soft-delete**: it updates every field mapped for the Delete action (the supplied mapping sets `INDACTIEF = 'N'`) and skips the update when the values already match. Delete fails when no fields are mapped for Delete and no stored procedure is configured. When the account does not exist, delete is skipped. The `WMS_GEBRCODE` record remains in the database.

**Delete stored procedure (optional):**

> [!WARNING]
> **Fit For Purpose (FFP)**
>
> `DeleteStoredProcedureName` is an optional, **Fit For Purpose (FFP)** extension: it was built for the specific requirements of one implementation that uses a customer-specific stored procedure. Such a procedure is not part of Key2 Belastingen itself, and **it is not certain that this approach applies to every implementation**. The consultant must review what the procedure does (for example which tables it cleans up and whether it deactivates `WMS_GEBRCODE`) and adjust the delete script where needed.

When a stored procedure is configured:

- Delete only executes `<procedure>('<GEBR_ORA>')` and reports success when the procedure does not raise an error; the field mapping for Delete is not applied
- The procedure is always executed, also for an already inactive account, because its effect is unknown to HelloID
- A stored procedure does not return its own logging to HelloID; this is stated in the audit log, and its logging must be checked in the database
- The connector does not verify the result of the procedure, so a change to the procedure does not require a change to the connector. `WMS_GEBRCODE` is re-queried afterwards only to return the current data
- Run this delete **before** the Key2BelastingenOracle delete, as such a procedure may require the Oracle user to still exist

### Menu-group permissions

Install `permissions/menuGroups` only where menu groups can be governed by HelloID Business Rules. This permissiontype is **Fit For Purpose (FFP)**, see [Fit For Purpose (FFP)](#fit-for-purpose-ffp).

- `permissions.ps1` lists all groups from `MENU_B_GROUP`, without filtering
- Grant checks the current assignment first to prevent duplicate rows (no known unique constraint on `MENU_B_USER`)
- Revoke deletes directly and logs a skip when no row was deleted
- `importPermissions.ps1` links `MENU_B_USER.USER_NAME` to `WMS_GEBRCODE.GEBR_ORA` and uses `GEBRCODE` as account reference, grouped per menu group in batches of 500. Only groups that exist in `MENU_B_GROUP` are imported, matching `permissions.ps1`
- Menu groups are not created or changed by this connector

### Function-specific rights

`MWB_GEBR_GROEP`, `MWB_STUUR`, `WMS_GEBRWVD` and `GHG_GEBRCODE` are not managed by this connector, because they require customer-specific mapping or copy logic. Assigning these rights remains a process outside HelloID (for example the customer's own scripts); consider a HelloID [Notification](https://docs.helloid.com/en/provisioning/notifications--provisioning-.html) after Create to inform the application manager.

### MDU integration

`ManageMduUser` is optional, disabled by default and **Fit For Purpose (FFP)**, see [Fit For Purpose (FFP)](#fit-for-purpose-ffp). When enabled, create inserts a `MDU_GEBRUIKER` record (`MDU_USER` = `GEBRCODE`, `MDU_IN_DIR` = lowercase Oracle username + `MduInDirectorySuffix`, `ORA_USER` = Oracle username) after the `WMS_GEBRCODE` record. Delete does not remove this record.

### Account import

`import.ps1` only imports records where `GEBR_ORA IS NOT NULL`; functional/process codes without an Oracle user are excluded in the script rather than through connector configuration. The imported fields are the import fields of the field mapping. An account is imported as enabled when `INDACTIEF` is `J`, and the `GEBR_ORA` is used as username.

### Governance reconciliation resolutions

Delete and Disable can be triggered from governance reconciliation (`$actionContext.ReconciliationOrigin = 'reconciliation'`), where no person context and no field mapping are available.

| Action | Values used in reconciliation mode |
| ------ | ---------------------------------- |
| Disable | `INDACTIEF = 'N'` |
| Delete (no stored procedure) | `INDACTIEF = 'N'` |
| Delete (stored procedure) | Only the account reference; the procedure is executed as in enforcement mode |

> [!IMPORTANT]
> The hardcoded reconciliation values are equal to the supplied field mapping. When the Disable or Delete field mapping is changed, update the values in `disable.ps1`/`delete.ps1` accordingly.

### Integrated security

When both Username and Password are empty, the DataSource must contain `Integrated Security=yes`; otherwise the action fails with a clear message. This only works when Oracle OS authentication is configured for the Windows account running the HelloID Agent. Configuring only one of Username or Password is rejected.

### Query execution

- Values are inserted as escaped SQL literals (single quotes doubled); table and column names come from the field mapping and must therefore be valid column names of `WMS_GEBRCODE`
- Statements are committed directly by Oracle and cannot be rolled back; a failed action can be partially executed (for example the `WMS_GEBRCODE` record is created but the optional `MDU_GEBRUIKER` insert fails) and the next run continues from the current state
- All actions that change something support DryRun; only the read queries are executed and the changes are logged. Import and the permissions list only read
- Each query uses a descriptive multiline variable and splatted parameters with an explicit `NonQuery` value

### Common Oracle errors

| Error | Action | Likely cause |
| ----- | ------ | ------------ |
| `ORA-00942` table or view does not exist | All actions | The service account lacks the `SELECT` right on the table, or the table name does not resolve in the schema of the service account. Check the [Requirements](#requirements) table |
| `ORA-01031` insufficient privileges | Create, update, enable, disable, delete, grant, revoke | The service account lacks an `INSERT`, `UPDATE`, `DELETE` or `EXECUTE` right from the [Requirements](#requirements) table |
| `ORA-20000: ##GHG-1005` | Create | The account was created inactive, see [Create behavior](#create-behavior) |
| `ORA-00001` unique constraint violated | Create | A record with the same `GEBRCODE` (or another unique column) already exists. Check the uniqueness configuration of `GEBRCODE` |
| `ORA-12154` / `ORA-12541` | All actions | The `DataSource` cannot be resolved or there is no listener; check the data source, port and network access of the agent |

## Development resources

### Database objects

The following database objects are used by the connector:

| Object | Used by | Description |
| ------ | ------- | ----------- |
| `WMS_GEBRCODE` | All account actions, uniqueness check, import, permissions | Key2 Belastingen application account |
| `MDU_GEBRUIKER` | Create (only when `ManageMduUser` is enabled) | Legacy MDU user record |
| `MENU_B_GROUP` | `permissions/menuGroups` | Available menu groups |
| `MENU_B_USER` | `permissions/menuGroups` | Menu-group assignments per Oracle user |
| `DeleteStoredProcedureName` | Delete (optional) | Customer-specific delete procedure, called with `GEBR_ORA` |

### Script conventions

- Each action builds a descriptive query variable (for example `$queryGetAccount` or `$queryUpdateAccount`) and invokes `Invoke-OracleQuery` through a matching splat. SELECT statements use `NonQuery = $false`; mutating statements use `NonQuery = $true`
- All actions use the same `New-OracleConnection`, `Invoke-OracleQuery` and `Resolve-OracleError` pattern as the Key2BelastingenOracle connector
- For an empty result, `Invoke-OracleQuery` returns a single `$null`; scripts therefore count with `Measure-Object` and filter with `Where-Object { $_ }` before `ForEach-Object`
- Oracle errors are detected via the base exception (`GetBaseException()`) and resolved by `Resolve-OracleError`

## Getting help

> [!TIP]
> For more information on how to configure a HelloID PowerShell connector, please refer to our [documentation](https://docs.helloid.com/en/provisioning/target-systems/powershell-v2-target-systems.html) pages.

## HelloID docs

The official HelloID documentation can be found at: https://docs.helloid.com/
