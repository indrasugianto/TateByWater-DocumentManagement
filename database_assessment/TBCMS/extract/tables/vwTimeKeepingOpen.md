# Table: vwTimeKeepingOpen *(linked)*

**Linked source:** `ODBC;DRIVER=SQL Server;SERVER=tbf-cms;APP=Microsoft Office;DATABASE=TateBywater` → `dbo.vwTimeKeepingOpen`

## Columns

| # | Column | Type | Size | Nullable | Default | Key | Notes |
|---|--------|------|------|----------|---------|-----|-------|
| 1 | Bill Closed | Boolean | 1 | yes | — | — | — |
| 2 | Bill_ID | Long | 4 | no | — | — | — |
| 3 | IANumber | Text | 255 | yes | — | — | allows zero-length string |
| 4 | FileNumber | Memo | — | yes | — | — | allows zero-length string |
| 5 | BalanceCalculated | Double | 8 | yes | — | — | — |
| 6 | CaseID | Long | 4 | no | — | PK | — |
| 7 | Last_Name | Text | 255 | yes | — | — | allows zero-length string |
| 8 | First_Name | Text | 255 | yes | — | — | allows zero-length string |
| 9 | CaseOpenDate | DateTime | 8 | yes | — | — | — |
| 10 | Closed | Boolean | 1 | yes | — | — | — |
| 11 | Clsdate | DateTime | 8 | yes | — | — | — |
| 12 | Extended_Ledger | Text | 255 | yes | — | — | allows zero-length string |
| 13 | Case_Letter | Text | 255 | yes | — | — | allows zero-length string |
| 14 | yr | Text | 254 | yes | — | — | allows zero-length string |
| 15 | Number_ | Long | 4 | yes | — | — | — |
| 16 | Orig_Atty | Text | 255 | yes | — | — | allows zero-length string |
| 17 | Address | Text | 255 | yes | — | — | allows zero-length string |
| 18 | CourtCaseNo | Text | 255 | yes | — | — | allows zero-length string |
| 19 | City | Text | 255 | yes | — | — | allows zero-length string |
| 20 | FamilyLaw | Boolean | 1 | yes | — | — | — |
| 21 | State | Text | 255 | yes | — | — | allows zero-length string |
| 22 | Zip | Text | 255 | yes | — | — | allows zero-length string |
| 23 | Country | Text | 255 | yes | — | — | allows zero-length string |
| 24 | HmPhone | Text | 255 | yes | — | — | allows zero-length string |
| 25 | Action | Text | 255 | yes | — | — | allows zero-length string |
| 26 | OtherPhone | Text | 255 | yes | — | — | allows zero-length string |
| 27 | Fax | Text | 255 | yes | — | — | allows zero-length string |
| 28 | WkPhone | Text | 255 | yes | — | — | allows zero-length string |
| 29 | Comments | Memo | — | yes | — | — | allows zero-length string |
| 30 | Email | Text | 255 | yes | — | — | allows zero-length string |
| 31 | Referral | Text | 255 | yes | — | — | allows zero-length string |
| 32 | Individual Referrer | Text | 255 | yes | — | — | allows zero-length string |
| 33 | Retainer | Currency | 8 | yes | — | — | — |
| 34 | Matter_type | Text | 255 | yes | — | — | allows zero-length string |
| 35 | SOL | DateTime | 8 | yes | — | — | — |
| 36 | Court | Text | 255 | yes | — | — | allows zero-length string |
| 37 | CType | Text | 255 | yes | — | — | allows zero-length string |
| 38 | POfc | Text | 255 | yes | — | — | allows zero-length string |
| 39 | ComplainingWitness | Text | 255 | yes | — | — | allows zero-length string |
| 40 | DOB | DateTime | 8 | yes | — | — | — |
| 41 | WkAddress | Text | 255 | yes | — | — | allows zero-length string |
| 42 | WkCity | Text | 255 | yes | — | — | allows zero-length string |
| 43 | WkState | Text | 255 | yes | — | — | allows zero-length string |
| 44 | WkZip | Text | 255 | yes | — | — | allows zero-length string |
| 45 | Pro Bono | Boolean | 1 | yes | — | — | — |
| 46 | HandlingAtty_Case | Text | 255 | yes | — | — | allows zero-length string |
| 47 | Action_Needed_on_Payment | Boolean | 1 | yes | — | — | — |
| 48 | SSN | Text | 255 | yes | — | — | allows zero-length string |
| 49 | Employer Name | Text | 255 | yes | — | — | allows zero-length string |
| 50 | Last Updated Contact Info | DateTime | 8 | yes | — | — | — |
| 51 | Ocounsel | Text | 255 | yes | — | — | allows zero-length string |
| 52 | Firm | Text | 255 | yes | — | — | allows zero-length string |
| 53 | OC Address | Text | 255 | yes | — | — | allows zero-length string |
| 54 | OC City | Text | 255 | yes | — | — | allows zero-length string |
| 55 | OC State | Text | 255 | yes | — | — | allows zero-length string |
| 56 | OC Zip | Text | 255 | yes | — | — | allows zero-length string |
| 57 | OC Phone | Text | 255 | yes | — | — | allows zero-length string |
| 58 | OC Email | Text | 255 | yes | — | — | allows zero-length string |
| 59 | OC Fax | Text | 255 | yes | — | — | allows zero-length string |
| 60 | Pro Bono PM | Text | 255 | yes | — | — | allows zero-length string |
| 61 | Pro Bono JRT | Text | 255 | yes | — | — | allows zero-length string |
| 62 | ContingencyFee | Boolean | 1 | yes | — | — | — |
| 63 | AuthorityToTalkTo | Memo | — | yes | — | — | allows zero-length string |
| 64 | Hourly | Boolean | 1 | yes | — | — | — |
| 65 | Contingency | Boolean | 1 | yes | — | — | — |
| 66 | Hybrid | Boolean | 1 | yes | — | — | — |
| 67 | Family-Law | Boolean | 1 | yes | — | — | — |
| 68 | Fixed | Boolean | 1 | yes | — | — | — |
| 69 | Scan | Boolean | 1 | yes | — | — | — |
| 70 | Scan Location | Memo | — | yes | — | — | allows zero-length string |
| 71 | ScanNotAvail | Boolean | 1 | yes | — | — | — |
| 72 | ParaLegal | Text | 255 | yes | — | — | allows zero-length string |
| 73 | Spanish | Boolean | 1 | yes | — | — | — |
| 74 | Offdate | DateTime | 8 | yes | — | — | — |
| 75 | CostHold | Currency | 8 | yes | — | — | — |
| 76 | CltNarrative | Memo | — | yes | — | — | allows zero-length string |
| 77 | ARTrustZero | Boolean | 1 | yes | — | — | — |
| 78 | F73 | Text | 255 | yes | — | — | allows zero-length string |
| 79 | F74 | Text | 255 | yes | — | — | allows zero-length string |
| 80 | F75 | Text | 255 | yes | — | — | allows zero-length string |
| 81 | F76 | Text | 255 | yes | — | — | allows zero-length string |
| 82 | PhName1 | Text | 255 | yes | — | — | allows zero-length string |
| 83 | PhName2 | Text | 255 | yes | — | — | allows zero-length string |
| 84 | LengthRes | Text | 255 | yes | — | — | allows zero-length string |
| 85 | LengthEmp | Text | 255 | yes | — | — | allows zero-length string |
| 86 | LegalStatus | Text | 255 | yes | — | — | allows zero-length string |
| 87 | CurrentBond | Text | 255 | yes | — | — | allows zero-length string |
| 88 | CrRecord | Memo | — | yes | — | — | allows zero-length string |
| 89 | TrustChronMemo | Memo | — | yes | — | — | allows zero-length string |
| 90 | Executor | Text | 255 | yes | — | — | allows zero-length string |
| 91 | RetainerReimb | Boolean | 1 | yes | — | — | — |
| 92 | RetReimbAmount | Currency | 8 | yes | — | — | — |
| 93 | Reviewable | Boolean | 1 | yes | — | — | — |
| 94 | ReviewReq | DateTime | 8 | yes | — | — | — |
| 95 | ReviewReceivedDate | DateTime | 8 | yes | — | — | — |
| 96 | ReviewReceived | Boolean | 1 | yes | — | — | — |
| 97 | Testimonial | Memo | — | yes | — | — | allows zero-length string |
| 98 | ReviewFollowUp | DateTime | 8 | yes | — | — | — |
| 99 | Stars | Long | 4 | yes | — | — | — |
| 100 | Review Source | Text | 255 | yes | — | — | allows zero-length string |
| 101 | Review Date | DateTime | 8 | yes | — | — | — |
| 102 | Title | Text | 255 | yes | — | — | allows zero-length string |
| 103 | OPartyLast | Text | 255 | yes | — | — | allows zero-length string |
| 104 | OPartyFirst | Text | 255 | yes | — | — | allows zero-length string |
| 105 | OPartyDOB | DateTime | 8 | yes | — | — | — |
| 106 | SSMA_TimeStamp | Binary | 8 | no | — | — | — |
| 107 | Bill Open | DateTime | 8 | yes | — | — | — |
| 108 | Name | Memo | — | no | — | — | allows zero-length string |

**Primary key:** CaseID

## Indexes

| Index | Fields | Primary | Unique | Foreign |
|-------|--------|---------|--------|---------|
| UniqueIndex | CaseID | yes | yes | no |

## Relationships

_No relationships declared in the database that reference this table._ Check column **lookup** notes above and query joins for implicit foreign keys.

---
*Generated by the extractor's `tableDefs` stage from `schema.json` + `relationships.json`. Structured source of truth: `schema.json`.*
