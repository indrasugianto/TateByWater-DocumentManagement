# Table: vwTimeKeepingClosed *(linked)*

**Linked source:** `ODBC;DRIVER=SQL Server;SERVER=awsql2022dev;APP=Microsoft Office;DATABASE=TateBywater` → `dbo.vwTimeKeepingClosed`

**Row count:** 95

## Columns

| # | Column | Type | Size | Nullable | Default | Key | Notes |
|---|--------|------|------|----------|---------|-----|-------|
| 1 | Bill Sent | DateTime | 8 | yes | — | — | — |
| 2 | Bill Paid | DateTime | 8 | yes | — | — | — |
| 3 | Bill Closed | Boolean | 1 | yes | — | — | — |
| 4 | BilL Closed Date | DateTime | 8 | yes | — | — | — |
| 5 | Discount | Currency | 8 | yes | — | — | — |
| 6 | IANumber | Text | 255 | yes | — | — | allows zero-length string |
| 7 | FileNumber | Memo | — | yes | — | — | allows zero-length string |
| 8 | BalanceCalculated | Double | 8 | yes | — | — | — |
| 9 | CaseID | Long | 4 | no | — | PK | — |
| 10 | Last_Name | Text | 255 | yes | — | — | allows zero-length string |
| 11 | First_Name | Text | 255 | yes | — | — | allows zero-length string |
| 12 | CaseOpenDate | DateTime | 8 | yes | — | — | — |
| 13 | Closed | Boolean | 1 | yes | — | — | — |
| 14 | Clsdate | DateTime | 8 | yes | — | — | — |
| 15 | Extended_Ledger | Text | 255 | yes | — | — | allows zero-length string |
| 16 | Case_Letter | Text | 255 | yes | — | — | allows zero-length string |
| 17 | yr | Text | 254 | yes | — | — | allows zero-length string |
| 18 | Number_ | Long | 4 | yes | — | — | — |
| 19 | Orig_Atty | Text | 255 | yes | — | — | allows zero-length string |
| 20 | Address | Text | 255 | yes | — | — | allows zero-length string |
| 21 | CourtCaseNo | Text | 255 | yes | — | — | allows zero-length string |
| 22 | City | Text | 255 | yes | — | — | allows zero-length string |
| 23 | FamilyLaw | Boolean | 1 | yes | — | — | — |
| 24 | State | Text | 255 | yes | — | — | allows zero-length string |
| 25 | Zip | Text | 255 | yes | — | — | allows zero-length string |
| 26 | Country | Text | 255 | yes | — | — | allows zero-length string |
| 27 | HmPhone | Text | 255 | yes | — | — | allows zero-length string |
| 28 | Action | Text | 255 | yes | — | — | allows zero-length string |
| 29 | OtherPhone | Text | 255 | yes | — | — | allows zero-length string |
| 30 | Fax | Text | 255 | yes | — | — | allows zero-length string |
| 31 | WkPhone | Text | 255 | yes | — | — | allows zero-length string |
| 32 | Comments | Memo | — | yes | — | — | allows zero-length string |
| 33 | Email | Text | 255 | yes | — | — | allows zero-length string |
| 34 | Referral | Text | 255 | yes | — | — | allows zero-length string |
| 35 | Individual Referrer | Text | 255 | yes | — | — | allows zero-length string |
| 36 | Retainer | Currency | 8 | yes | — | — | — |
| 37 | Matter_type | Text | 255 | yes | — | — | allows zero-length string |
| 38 | SOL | DateTime | 8 | yes | — | — | — |
| 39 | Court | Text | 255 | yes | — | — | allows zero-length string |
| 40 | CType | Text | 255 | yes | — | — | allows zero-length string |
| 41 | POfc | Text | 255 | yes | — | — | allows zero-length string |
| 42 | ComplainingWitness | Text | 255 | yes | — | — | allows zero-length string |
| 43 | DOB | DateTime | 8 | yes | — | — | — |
| 44 | WkAddress | Text | 255 | yes | — | — | allows zero-length string |
| 45 | WkCity | Text | 255 | yes | — | — | allows zero-length string |
| 46 | WkState | Text | 255 | yes | — | — | allows zero-length string |
| 47 | WkZip | Text | 255 | yes | — | — | allows zero-length string |
| 48 | Pro Bono | Boolean | 1 | yes | — | — | — |
| 49 | HandlingAtty_Case | Text | 255 | yes | — | — | allows zero-length string |
| 50 | Action_Needed_on_Payment | Boolean | 1 | yes | — | — | — |
| 51 | SSN | Text | 255 | yes | — | — | allows zero-length string |
| 52 | Employer Name | Text | 255 | yes | — | — | allows zero-length string |
| 53 | Last Updated Contact Info | DateTime | 8 | yes | — | — | — |
| 54 | Ocounsel | Text | 255 | yes | — | — | allows zero-length string |
| 55 | Firm | Text | 255 | yes | — | — | allows zero-length string |
| 56 | OC Address | Text | 255 | yes | — | — | allows zero-length string |
| 57 | OC City | Text | 255 | yes | — | — | allows zero-length string |
| 58 | OC State | Text | 255 | yes | — | — | allows zero-length string |
| 59 | OC Zip | Text | 255 | yes | — | — | allows zero-length string |
| 60 | OC Phone | Text | 255 | yes | — | — | allows zero-length string |
| 61 | OC Email | Text | 255 | yes | — | — | allows zero-length string |
| 62 | OC Fax | Text | 255 | yes | — | — | allows zero-length string |
| 63 | Pro Bono PM | Text | 255 | yes | — | — | allows zero-length string |
| 64 | Pro Bono JRT | Text | 255 | yes | — | — | allows zero-length string |
| 65 | ContingencyFee | Boolean | 1 | yes | — | — | — |
| 66 | AuthorityToTalkTo | Memo | — | yes | — | — | allows zero-length string |
| 67 | Hourly | Boolean | 1 | yes | — | — | — |
| 68 | Contingency | Boolean | 1 | yes | — | — | — |
| 69 | Hybrid | Boolean | 1 | yes | — | — | — |
| 70 | Family-Law | Boolean | 1 | yes | — | — | — |
| 71 | Fixed | Boolean | 1 | yes | — | — | — |
| 72 | Scan | Boolean | 1 | yes | — | — | — |
| 73 | Scan Location | Memo | — | yes | — | — | allows zero-length string |
| 74 | ScanNotAvail | Boolean | 1 | yes | — | — | — |
| 75 | ParaLegal | Text | 255 | yes | — | — | allows zero-length string |
| 76 | Spanish | Boolean | 1 | yes | — | — | — |
| 77 | Offdate | DateTime | 8 | yes | — | — | — |
| 78 | CostHold | Currency | 8 | yes | — | — | — |
| 79 | CltNarrative | Memo | — | yes | — | — | allows zero-length string |
| 80 | ARTrustZero | Boolean | 1 | yes | — | — | — |
| 81 | F73 | Text | 255 | yes | — | — | allows zero-length string |
| 82 | F74 | Text | 255 | yes | — | — | allows zero-length string |
| 83 | F75 | Text | 255 | yes | — | — | allows zero-length string |
| 84 | F76 | Text | 255 | yes | — | — | allows zero-length string |
| 85 | PhName1 | Text | 255 | yes | — | — | allows zero-length string |
| 86 | PhName2 | Text | 255 | yes | — | — | allows zero-length string |
| 87 | LengthRes | Text | 255 | yes | — | — | allows zero-length string |
| 88 | LengthEmp | Text | 255 | yes | — | — | allows zero-length string |
| 89 | LegalStatus | Text | 255 | yes | — | — | allows zero-length string |
| 90 | CurrentBond | Text | 255 | yes | — | — | allows zero-length string |
| 91 | CrRecord | Memo | — | yes | — | — | allows zero-length string |
| 92 | TrustChronMemo | Memo | — | yes | — | — | allows zero-length string |
| 93 | Executor | Text | 255 | yes | — | — | allows zero-length string |
| 94 | RetainerReimb | Boolean | 1 | yes | — | — | — |
| 95 | RetReimbAmount | Currency | 8 | yes | — | — | — |
| 96 | Reviewable | Boolean | 1 | yes | — | — | — |
| 97 | ReviewReq | DateTime | 8 | yes | — | — | — |
| 98 | ReviewReceivedDate | DateTime | 8 | yes | — | — | — |
| 99 | ReviewReceived | Boolean | 1 | yes | — | — | — |
| 100 | Testimonial | Memo | — | yes | — | — | allows zero-length string |
| 101 | ReviewFollowUp | DateTime | 8 | yes | — | — | — |
| 102 | Stars | Long | 4 | yes | — | — | — |
| 103 | Review Source | Text | 255 | yes | — | — | allows zero-length string |
| 104 | Review Date | DateTime | 8 | yes | — | — | — |
| 105 | Title | Text | 255 | yes | — | — | allows zero-length string |
| 106 | OPartyLast | Text | 255 | yes | — | — | allows zero-length string |
| 107 | OPartyFirst | Text | 255 | yes | — | — | allows zero-length string |
| 108 | OPartyDOB | DateTime | 8 | yes | — | — | — |
| 109 | SSMA_TimeStamp | Binary | 8 | no | — | — | — |
| 110 | Bill Open | DateTime | 8 | yes | — | — | — |
| 111 | Name | Memo | — | no | — | — | allows zero-length string |
| 112 | Bill_ID | Long | 4 | no | — | — | — |
| 113 | TrustatClose | Currency | 8 | yes | — | — | — |
| 114 | StatementLessTrust | Boolean | 1 | yes | — | — | — |
| 115 | InvoiceExceedsTrust | Boolean | 1 | yes | — | — | — |
| 116 | InvoiceTotalAdvance | Boolean | 1 | yes | — | — | — |
| 117 | InvoiceNoAdvance | Boolean | 1 | yes | — | — | — |

**Primary key:** CaseID

## Indexes

| Index | Fields | Primary | Unique | Foreign |
|-------|--------|---------|--------|---------|
| UniqueIndex | CaseID | yes | yes | no |

## Relationships

_No relationships declared in the database that reference this table._ Check column **lookup** notes above and query joins for implicit foreign keys.

---
*Generated by the extractor's `tableDefs` stage from `schema.json` + `relationships.json`. Structured source of truth: `schema.json`.*
