# Reference: HAT output contract used by HOLDING_79_TRANSFER

This document freezes only the **external XLSX shape** reused from the owner-provided HAT handoff.
HAT business logic, matching rules, hierarchy transfer algorithm, and account semantics are **not** reused.

## Workbook boundary

- one workbook = one document organization + one document date;
- the financial sheet is named exactly `Загрузка_A_AA`;
- the sheet contains exactly 27 columns in this order;
- no service/debug columns may be inserted into this sheet;
- internal audit/control data belongs in separate run artifacts, not in the loader sheet.

## Exact 27-column order

1. `СчетДт`
2. `СчетКт`
3. `ВалютаДт`
4. `ВалютаКт`
5. `ВидОперации`
6. `ПодразделениеДт`
7. `ПодразделениеКт`
8. `НаправлениеДеятельностиДт`
9. `НаправлениеДеятельностиКт`
10. `СуммаВВалютеУчета`
11. `СуммаВВалютеОтчетности`
12. `СуммаВВалютеДт`
13. `СуммаВВалютеКт`
14. `КоличествоДт`
15. `КоличествоКт`
16. `Содержание`
17. `СчетДтИсточник`
18. `СчетКтИсточник`
19. `ИдентификаторФинЗаписи`
20. `ПравилоДт`
21. `ПравилоКт`
22. `СубконтоДт1`
23. `СубконтоДт2`
24. `СубконтоДт3`
25. `СубконтоКт1`
26. `СубконтоКт2`
27. `СубконтоКт3`

## HOLDING_79_TRANSFER ownership of values

The new project must define its own deterministic mapping from canonical `PostingRow` to these columns.
For accounts `79.1`, `79.2`, `79.3`, `Поставщик РВП` is represented by the project's approved subkonto mapping and department by `ПодразделениеДт/Кт`.

Unknown fields must never be guessed. Until an exact rule is accepted, leave the field empty only if the loader contract permits it; otherwise block export with an explicit reason.

`ВидОперации`, `Содержание`, source-account fields, rule fields, financial-record id, currency fields, and unused subkonto slots are implementation-contract decisions for Issue #1 and must be covered by tests before exporter work starts.
