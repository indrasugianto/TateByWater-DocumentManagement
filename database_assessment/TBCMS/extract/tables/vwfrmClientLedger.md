# Table: vwfrmClientLedger *(linked)*

**Linked source:** `ODBC;DRIVER=SQL Server;SERVER=tbf-cms;APP=Microsoft Office;DATABASE=TateBywater` → `dbo.vwfrmClientLedger`

## Columns

| # | Column | Type | Size | Nullable | Default | Key | Notes |
|---|--------|------|------|----------|---------|-----|-------|
| 1 | CaseID | Long | 4 | no | — | PK | auto-increment |
| 2 | Last_Name | Text | 255 | yes | — | — | allows zero-length string |
| 3 | First_Name | Text | 255 | yes | — | — | allows zero-length string |
| 4 | CaseOpenDate | DateTime | 8 | yes | — | — | — |
| 5 | Closed | Boolean | 1 | yes | — | — | — |
| 6 | Clsdate | DateTime | 8 | yes | — | — | — |
| 7 | Extended_Ledger | Text | 255 | yes | — | — | allows zero-length string |
| 8 | Case_Letter | Text | 255 | yes | — | — | allows zero-length string |
| 9 | yr | Text | 254 | yes | — | — | allows zero-length string |
| 10 | Number_ | Long | 4 | yes | — | — | — |
| 11 | Orig_Atty | Text | 255 | yes | — | — | allows zero-length string |
| 12 | Address | Text | 255 | yes | — | — | allows zero-length string |
| 13 | CourtCaseNo | Text | 255 | yes | — | — | allows zero-length string |
| 14 | City | Text | 255 | yes | — | — | allows zero-length string |
| 15 | FamilyLaw | Boolean | 1 | yes | — | — | — |
| 16 | State | Text | 255 | yes | — | — | allows zero-length string |
| 17 | Zip | Text | 255 | yes | — | — | allows zero-length string |
| 18 | Country | Text | 255 | yes | — | — | allows zero-length string |
| 19 | HmPhone | Text | 255 | yes | — | — | allows zero-length string |
| 20 | Action | Text | 255 | yes | — | — | allows zero-length string |
| 21 | OtherPhone | Text | 255 | yes | — | — | allows zero-length string |
| 22 | Fax | Text | 255 | yes | — | — | allows zero-length string |
| 23 | WkPhone | Text | 255 | yes | — | — | allows zero-length string |
| 24 | Comments | Memo | — | yes | — | — | allows zero-length string |
| 25 | Email | Text | 255 | yes | — | — | allows zero-length string |
| 26 | Referral | Text | 255 | yes | — | — | allows zero-length string |
| 27 | Individual Referrer | Text | 255 | yes | — | — | allows zero-length string |
| 28 | Retainer | Currency | 8 | yes | — | — | — |
| 29 | Matter_type | Text | 255 | yes | — | — | allows zero-length string |
| 30 | SOL | DateTime | 8 | yes | — | — | — |
| 31 | Court | Text | 255 | yes | — | — | allows zero-length string |
| 32 | CType | Text | 255 | yes | — | — | allows zero-length string |
| 33 | POfc | Text | 255 | yes | — | — | allows zero-length string |
| 34 | ComplainingWitness | Text | 255 | yes | — | — | allows zero-length string |
| 35 | DOB | DateTime | 8 | yes | — | — | — |
| 36 | WkAddress | Text | 255 | yes | — | — | allows zero-length string |
| 37 | WkCity | Text | 255 | yes | — | — | allows zero-length string |
| 38 | WkState | Text | 255 | yes | — | — | allows zero-length string |
| 39 | WkZip | Text | 255 | yes | — | — | allows zero-length string |
| 40 | Pro Bono | Boolean | 1 | yes | — | — | — |
| 41 | HandlingAtty_Case | Text | 255 | yes | — | — | allows zero-length string |
| 42 | Action_Needed_on_Payment | Boolean | 1 | yes | — | — | — |
| 43 | SSN | Text | 255 | yes | — | — | allows zero-length string |
| 44 | Employer Name | Text | 255 | yes | — | — | allows zero-length string |
| 45 | Last Updated Contact Info | DateTime | 8 | yes | — | — | — |
| 46 | Ocounsel | Text | 255 | yes | — | — | allows zero-length string |
| 47 | Firm | Text | 255 | yes | — | — | allows zero-length string |
| 48 | OC Address | Text | 255 | yes | — | — | allows zero-length string |
| 49 | OC City | Text | 255 | yes | — | — | allows zero-length string |
| 50 | OC State | Text | 255 | yes | — | — | allows zero-length string |
| 51 | OC Zip | Text | 255 | yes | — | — | allows zero-length string |
| 52 | OC Phone | Text | 255 | yes | — | — | allows zero-length string |
| 53 | OC Email | Text | 255 | yes | — | — | allows zero-length string |
| 54 | OC Fax | Text | 255 | yes | — | — | allows zero-length string |
| 55 | Pro Bono PM | Text | 255 | yes | — | — | allows zero-length string |
| 56 | Pro Bono JRT | Text | 255 | yes | — | — | allows zero-length string |
| 57 | ContingencyFee | Boolean | 1 | yes | — | — | — |
| 58 | AuthorityToTalkTo | Memo | — | yes | — | — | allows zero-length string |
| 59 | Hourly | Boolean | 1 | yes | — | — | — |
| 60 | Contingency | Boolean | 1 | yes | — | — | — |
| 61 | Hybrid | Boolean | 1 | yes | — | — | — |
| 62 | Family-Law | Boolean | 1 | yes | — | — | — |
| 63 | Fixed | Boolean | 1 | yes | — | — | — |
| 64 | Scan | Boolean | 1 | yes | — | — | — |
| 65 | Scan Location | Memo | — | yes | — | — | allows zero-length string |
| 66 | ScanNotAvail | Boolean | 1 | yes | — | — | — |
| 67 | ParaLegal | Text | 255 | yes | — | — | allows zero-length string |
| 68 | Spanish | Boolean | 1 | yes | — | — | — |
| 69 | Offdate | DateTime | 8 | yes | — | — | — |
| 70 | CostHold | Currency | 8 | yes | — | — | — |
| 71 | CltNarrative | Memo | — | yes | — | — | allows zero-length string |
| 72 | ARTrustZero | Boolean | 1 | yes | — | — | — |
| 73 | F73 | Text | 255 | yes | — | — | allows zero-length string |
| 74 | F74 | Text | 255 | yes | — | — | allows zero-length string |
| 75 | F75 | Text | 255 | yes | — | — | allows zero-length string |
| 76 | F76 | Text | 255 | yes | — | — | allows zero-length string |
| 77 | PhName1 | Text | 255 | yes | — | — | allows zero-length string |
| 78 | PhName2 | Text | 255 | yes | — | — | allows zero-length string |
| 79 | LengthRes | Text | 255 | yes | — | — | allows zero-length string |
| 80 | LengthEmp | Text | 255 | yes | — | — | allows zero-length string |
| 81 | LegalStatus | Text | 255 | yes | — | — | allows zero-length string |
| 82 | CurrentBond | Text | 255 | yes | — | — | allows zero-length string |
| 83 | CrRecord | Memo | — | yes | — | — | allows zero-length string |
| 84 | TrustChronMemo | Memo | — | yes | — | — | allows zero-length string |
| 85 | Executor | Text | 255 | yes | — | — | allows zero-length string |
| 86 | RetainerReimb | Boolean | 1 | yes | — | — | — |
| 87 | RetReimbAmount | Currency | 8 | yes | — | — | — |
| 88 | Reviewable | Boolean | 1 | yes | — | — | — |
| 89 | ReviewReq | DateTime | 8 | yes | — | — | — |
| 90 | ReviewReceivedDate | DateTime | 8 | yes | — | — | — |
| 91 | ReviewReceived | Boolean | 1 | yes | — | — | — |
| 92 | Testimonial | Memo | — | yes | — | — | allows zero-length string |
| 93 | ReviewFollowUp | DateTime | 8 | yes | — | — | — |
| 94 | Stars | Long | 4 | yes | — | — | — |
| 95 | Review Source | Text | 255 | yes | — | — | allows zero-length string |
| 96 | Review Date | DateTime | 8 | yes | — | — | — |
| 97 | Title | Text | 255 | yes | — | — | allows zero-length string |
| 98 | OPartyLast | Text | 255 | yes | — | — | allows zero-length string |
| 99 | OPartyFirst | Text | 255 | yes | — | — | allows zero-length string |
| 100 | OPartyDOB | DateTime | 8 | yes | — | — | — |
| 101 | SSMA_TimeStamp | Binary | 8 | no | — | — | — |
| 102 | FileNo | Memo | — | yes | — | — | allows zero-length string |
| 103 | PartnerRate | Currency | 8 | yes | — | — | — |
| 104 | AssocRate | Currency | 8 | yes | — | — | — |
| 105 | FileLocation | Text | 50 | yes | — | — | allows zero-length string |

**Primary key:** CaseID

## Indexes

| Index | Fields | Primary | Unique | Foreign |
|-------|--------|---------|--------|---------|
| UniqueIndex | CaseID | yes | yes | no |

## Relationships

_No relationships declared in the database that reference this table._ Check column **lookup** notes above and query joins for implicit foreign keys.

---
*Generated by the extractor's `tableDefs` stage from `schema.json` + `relationships.json`. Structured source of truth: `schema.json`.*
