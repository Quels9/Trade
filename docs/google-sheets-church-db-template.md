# Google Sheets Church Database Template

Use one Google Spreadsheet with the following tabs in this exact order:

1. `Members`
2. `Families`
3. `Attendance`
4. `Giving`
5. `Notes`
6. `Lookups`

## 1) Members tab

### Headers (row 1)
```
member_id | first_name | last_name | full_name | phone | email | address | city | state | zip | birthday | join_date | family_id | status | ministry_group | last_attendance_date | notes_count | created_at | updated_at
```

### Formulas
- `full_name` (D2):
```gs
=ARRAYFORMULA(IF(A2:A="","",B2:B&" "&C2:C))
```
- `notes_count` (Q2):
```gs
=ARRAYFORMULA(IF(A2:A="","",COUNTIF(Notes!B:B,A2:A)))
```

### Validations
- `status` (N:N): List from range `Lookups!A2:A` (example values: Active, Inactive, Visitor).
- `ministry_group` (O:O): List from range `Lookups!B2:B`.

## 2) Families tab

### Headers
```
family_id | family_name | address | city | state | zip | primary_contact_member_id | phone | email | created_at | updated_at
```

## 3) Attendance tab

### Headers
```
attendance_id | date | service_type | member_id | present | checkin_source | created_at
```

### Validations
- `service_type` (C:C): List from range `Lookups!C2:C`.
- `present` (E:E): Checkbox.
- `checkin_source` (F:F): List from range `Lookups!D2:D`.

## 4) Giving tab (sensitive; lock this sheet)

### Headers
```
gift_id | date | member_id | amount | method | fund | receipt_sent | created_at
```

### Validations
- `method` (E:E): List from range `Lookups!E2:E`.
- `fund` (F:F): List from range `Lookups!F2:F`.
- `receipt_sent` (G:G): Checkbox.

## 5) Notes tab

### Headers
```
note_id | member_id | note_type | note_text | followup_date | created_by | created_at
```

### Validations
- `note_type` (C:C): List from range `Lookups!G2:G`.

## 6) Lookups tab

Place these lookup lists:

- Column A (`status`): `Active`, `Inactive`, `Visitor`
- Column B (`ministry_group`): `Kids`, `Youth`, `Worship`, `Outreach`, `Small Group`
- Column C (`service_type`): `Sunday AM`, `Sunday PM`, `Midweek`, `Special Event`
- Column D (`checkin_source`): `Form`, `Manual`, `Import`
- Column E (`giving_method`): `Cash`, `Check`, `Card`, `Online`, `Zelle`
- Column F (`fund`): `Tithes`, `Missions`, `Building`, `Benevolence`, `Other`
- Column G (`note_type`): `Pastoral`, `Care`, `Prayer`, `Follow-up`

## Conditional formatting

1. `Members!N:N` where value = `Inactive` -> light red background.
2. `Attendance!B:B` where date is older than 30 days and row has member_id -> light yellow.
3. `Giving!D:D` where amount > 1000 -> light green.

## Suggested protections

- Protect `Members` formula columns (`D`, `Q`, `R`, `S`).
- Protect entire `Giving` and `Notes` tabs except for finance/pastoral roles.
- Restrict editors to trusted admins.

## Automation hooks expected by Apps Script

The provided Apps Script expects these columns to exist and will auto-populate:
- IDs in column A for `Members`, `Families`, `Attendance`, `Giving`, `Notes`.
- timestamps in `created_at` and `updated_at` columns where present.
- duplicate email/phone checks on `Members`.
- `last_attendance_date` in `Members` from attendance updates.
