# Master Pricelist AC

## ERD

```mermaid
erDiagram
    CUSTOMER ||--o{ QUOTATION : receives
    CUSTOMER ||--o{ MATERIAL_ESTIMATE : requests
    AC_BRAND ||--o{ AC_PRODUCT : owns
    AC_PRODUCT ||--o{ INSTALLATION_PACKAGE_ITEM : contains
    MATERIAL ||--o{ INSTALLATION_PACKAGE_ITEM : contains
    CATALOG_SERVICE ||--o{ INSTALLATION_PACKAGE_ITEM : contains
    INSTALLATION_PACKAGE ||--o{ INSTALLATION_PACKAGE_ITEM : has
    QUOTATION ||--o{ QUOTATION_ITEM : contains
    AC_PRODUCT ||--o{ QUOTATION_ITEM : priced_as_snapshot
    MATERIAL ||--o{ QUOTATION_ITEM : priced_as_snapshot
    CATALOG_SERVICE ||--o{ QUOTATION_ITEM : priced_as_snapshot
    QUOTATION ||--o| MATERIAL_ESTIMATE : informs
    MATERIAL_ESTIMATE ||--o{ MATERIAL_ESTIMATE_ITEM : contains
    MATERIAL ||--o{ MATERIAL_ESTIMATE_ITEM : priced_as_snapshot
    CATALOG_SERVICE ||--o{ MATERIAL_ESTIMATE_ITEM : priced_as_snapshot
    CUSTOMER ||--o{ INVOICE : billed
```

## Transaction Rules

- `QuotationItem` and `MaterialEstimateItem` retain cost and sell-price snapshots. Updating a master record does not alter a historic quotation, invoice conversion, profit report, or estimate.
- A quotation is converted to an existing `Invoice` and `InvoiceItem` structure with the quotation number in `Invoice.spo`.
- Inactive masters are excluded from selector APIs but retained for historical references.
- Apply the migration with `flask --app app db upgrade`, then use **Seed Master Indonesia** on the Master Pricelist page for brands, common services/materials, and the Gree F5S example.

## REST API

| Method | Endpoint | Purpose |
| --- | --- | --- |
| GET/POST | `/api/master-pricelist/products` | List/create AC products |
| GET/POST | `/api/master-pricelist/materials` | List/create materials |
| GET/POST | `/api/master-pricelist/services` | List/create services |
| GET/POST | `/api/master-pricelist/packages` | List/create installation packages |
| PATCH/DELETE | `/api/master-pricelist/<resource>/<id>` | Update / soft deactivate master data |
| POST | `/api/master-pricelist/load-calculator` | Calculate electrical load |