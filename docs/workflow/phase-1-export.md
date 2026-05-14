# Phase 1 — Export a Selection Workbook

## Steps

1. Run the tool (double-click `run.bat`).
2. Choose option **[1] Export policies to an Excel selection file**.
3. Pick one or more tenants.
4. For each tenant, the tool writes `output\<tenant>_Policy_Selection.xlsx`.

## What's in the workbook

The workbook has one row per tenant policy, with these columns:

| Column | Meaning |
|--------|---------|
| **Include** | `Yes` / `No` dropdown. Drives what ends up in the SOP. Policies that look like they already match a canonical SOP control are pre-ticked `Yes`. |
| **Policy Name** | Friendly name as reported by Inforcer. |
| **Product** | Product area (e.g. Entra, Intune). |
| **Primary / Secondary Group** | Inforcer category metadata. |
| **Suggested Section** | The Part 2 section the policy will land in (2.1 – 2.7). |
| **Matches Template Control** | Canonical controls this policy would populate. |
| **Settings Summary** | Compact JSON-ish view of the policy's current settings — so you can eyeball what would be inserted into the SOP. |
| **Policy ID** | Internal ID used to re-fetch the policy in Phase 3. Leave it alone. |

## Tenant metadata rows

!!! warning "Don't edit the top rows"
    Rows 1–3 carry the tenant metadata (name, Inforcer ID, region). Don't edit those — they're how Phase 3 knows which tenant this workbook belongs to.

## Next step

Move on to [Phase 2 — Review](phase-2-review.md).
