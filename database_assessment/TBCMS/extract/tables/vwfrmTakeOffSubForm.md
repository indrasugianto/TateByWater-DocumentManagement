# Table: vwfrmTakeOffSubForm *(linked)*

**Linked source:** `ODBC;DRIVER=SQL Server;SERVER=awsql2022dev;APP=Microsoft Office;DATABASE=TateBywater` → `dbo.vwfrmTakeOffSubForm`

**Row count:** 52463

## Columns

| # | Column | Type | Size | Nullable | Default | Key | Notes |
|---|--------|------|------|----------|---------|-----|-------|
| 1 | FileNumber | Memo | — | yes | — | — | allows zero-length string |
| 2 | Name | Memo | — | no | — | — | allows zero-length string |
| 3 | CaseID | Long | 4 | no | — | — | — |
| 4 | Last_Name | Text | 255 | yes | — | — | allows zero-length string |
| 5 | First_Name | Text | 255 | yes | — | — | allows zero-length string |
| 6 | CaseOpenDate | DateTime | 8 | yes | — | — | — |
| 7 | Closed | Boolean | 1 | yes | — | — | — |
| 8 | Clsdate | DateTime | 8 | yes | — | — | — |
| 9 | Extended_Ledger | Text | 255 | yes | — | — | allows zero-length string |
| 10 | Case_Letter | Text | 255 | yes | — | — | allows zero-length string |
| 11 | yr | Text | 254 | yes | — | — | allows zero-length string |
| 12 | Number_ | Long | 4 | yes | — | — | — |
| 13 | Orig_Atty | Text | 255 | yes | — | — | allows zero-length string |
| 14 | Address | Text | 255 | yes | — | — | allows zero-length string |
| 15 | CourtCaseNo | Text | 255 | yes | — | — | allows zero-length string |
| 16 | City | Text | 255 | yes | — | — | allows zero-length string |
| 17 | FamilyLaw | Boolean | 1 | yes | — | — | — |
| 18 | State | Text | 255 | yes | — | — | allows zero-length string |
| 19 | Zip | Text | 255 | yes | — | — | allows zero-length string |
| 20 | Country | Text | 255 | yes | — | — | allows zero-length string |
| 21 | HmPhone | Text | 255 | yes | — | — | allows zero-length string |
| 22 | Action | Text | 255 | yes | — | — | allows zero-length string |
| 23 | OtherPhone | Text | 255 | yes | — | — | allows zero-length string |
| 24 | Fax | Text | 255 | yes | — | — | allows zero-length string |
| 25 | WkPhone | Text | 255 | yes | — | — | allows zero-length string |
| 26 | Comments | Memo | — | yes | — | — | allows zero-length string |
| 27 | Email | Text | 255 | yes | — | — | allows zero-length string |
| 28 | Referral | Text | 255 | yes | — | — | allows zero-length string |
| 29 | Individual Referrer | Text | 255 | yes | — | — | allows zero-length string |
| 30 | Retainer | Currency | 8 | yes | — | — | — |
| 31 | Matter_type | Text | 255 | yes | — | — | allows zero-length string |
| 32 | SOL | DateTime | 8 | yes | — | — | — |
| 33 | Court | Text | 255 | yes | — | — | allows zero-length string |
| 34 | CType | Text | 255 | yes | — | — | allows zero-length string |
| 35 | POfc | Text | 255 | yes | — | — | allows zero-length string |
| 36 | ComplainingWitness | Text | 255 | yes | — | — | allows zero-length string |
| 37 | DOB | DateTime | 8 | yes | — | — | — |
| 38 | WkAddress | Text | 255 | yes | — | — | allows zero-length string |
| 39 | WkCity | Text | 255 | yes | — | — | allows zero-length string |
| 40 | WkState | Text | 255 | yes | — | — | allows zero-length string |
| 41 | WkZip | Text | 255 | yes | — | — | allows zero-length string |
| 42 | Pro Bono | Boolean | 1 | yes | — | — | — |
| 43 | HandlingAtty_Case | Text | 255 | yes | — | — | allows zero-length string |
| 44 | Action_Needed_on_Payment | Boolean | 1 | yes | — | — | — |
| 45 | SSN | Text | 255 | yes | — | — | allows zero-length string |
| 46 | Employer Name | Text | 255 | yes | — | — | allows zero-length string |
| 47 | Last Updated Contact Info | DateTime | 8 | yes | — | — | — |
| 48 | Ocounsel | Text | 255 | yes | — | — | allows zero-length string |
| 49 | Firm | Text | 255 | yes | — | — | allows zero-length string |
| 50 | OC Address | Text | 255 | yes | — | — | allows zero-length string |
| 51 | OC City | Text | 255 | yes | — | — | allows zero-length string |
| 52 | OC State | Text | 255 | yes | — | — | allows zero-length string |
| 53 | OC Zip | Text | 255 | yes | — | — | allows zero-length string |
| 54 | OC Phone | Text | 255 | yes | — | — | allows zero-length string |
| 55 | OC Email | Text | 255 | yes | — | — | allows zero-length string |
| 56 | OC Fax | Text | 255 | yes | — | — | allows zero-length string |
| 57 | Pro Bono PM | Text | 255 | yes | — | — | allows zero-length string |
| 58 | Pro Bono JRT | Text | 255 | yes | — | — | allows zero-length string |
| 59 | ContingencyFee | Boolean | 1 | yes | — | — | — |
| 60 | AuthorityToTalkTo | Memo | — | yes | — | — | allows zero-length string |
| 61 | Hourly | Boolean | 1 | yes | — | — | — |
| 62 | Contingency | Boolean | 1 | yes | — | — | — |
| 63 | Hybrid | Boolean | 1 | yes | — | — | — |
| 64 | Family-Law | Boolean | 1 | yes | — | — | — |
| 65 | Fixed | Boolean | 1 | yes | — | — | — |
| 66 | Scan | Boolean | 1 | yes | — | — | — |
| 67 | Scan Location | Memo | — | yes | — | — | allows zero-length string |
| 68 | ScanNotAvail | Boolean | 1 | yes | — | — | — |
| 69 | ParaLegal | Text | 255 | yes | — | — | allows zero-length string |
| 70 | Spanish | Boolean | 1 | yes | — | — | — |
| 71 | Offdate | DateTime | 8 | yes | — | — | — |
| 72 | CostHold | Currency | 8 | yes | — | — | — |
| 73 | CltNarrative | Memo | — | yes | — | — | allows zero-length string |
| 74 | ARTrustZero | Boolean | 1 | yes | — | — | — |
| 75 | F73 | Text | 255 | yes | — | — | allows zero-length string |
| 76 | F74 | Text | 255 | yes | — | — | allows zero-length string |
| 77 | F75 | Text | 255 | yes | — | — | allows zero-length string |
| 78 | F76 | Text | 255 | yes | — | — | allows zero-length string |
| 79 | PhName1 | Text | 255 | yes | — | — | allows zero-length string |
| 80 | PhName2 | Text | 255 | yes | — | — | allows zero-length string |
| 81 | LengthRes | Text | 255 | yes | — | — | allows zero-length string |
| 82 | LengthEmp | Text | 255 | yes | — | — | allows zero-length string |
| 83 | LegalStatus | Text | 255 | yes | — | — | allows zero-length string |
| 84 | CurrentBond | Text | 255 | yes | — | — | allows zero-length string |
| 85 | CrRecord | Memo | — | yes | — | — | allows zero-length string |
| 86 | TrustChronMemo | Memo | — | yes | — | — | allows zero-length string |
| 87 | Executor | Text | 255 | yes | — | — | allows zero-length string |
| 88 | RetainerReimb | Boolean | 1 | yes | — | — | — |
| 89 | RetReimbAmount | Currency | 8 | yes | — | — | — |
| 90 | Reviewable | Boolean | 1 | yes | — | — | — |
| 91 | ReviewReq | DateTime | 8 | yes | — | — | — |
| 92 | ReviewReceivedDate | DateTime | 8 | yes | — | — | — |
| 93 | ReviewReceived | Boolean | 1 | yes | — | — | — |
| 94 | Testimonial | Memo | — | yes | — | — | allows zero-length string |
| 95 | ReviewFollowUp | DateTime | 8 | yes | — | — | — |
| 96 | Stars | Long | 4 | yes | — | — | — |
| 97 | Review Source | Text | 255 | yes | — | — | allows zero-length string |
| 98 | Review Date | DateTime | 8 | yes | — | — | — |
| 99 | Title | Text | 255 | yes | — | — | allows zero-length string |
| 100 | OPartyLast | Text | 255 | yes | — | — | allows zero-length string |
| 101 | OPartyFirst | Text | 255 | yes | — | — | allows zero-length string |
| 102 | OPartyDOB | DateTime | 8 | yes | — | — | — |
| 103 | TakeOffID | Long | 4 | yes | — | PK | — |
| 104 | TakeOffMonthID | Long | 4 | yes | — | — | — |
| 105 | AvailBalance | Currency | 8 | yes | — | — | — |
| 106 | TotalUnCashedChks | Currency | 8 | yes | — | — | — |
| 107 | TotalUnclearedDeps | Currency | 8 | yes | — | — | — |
| 108 | TotalAdvancedAR | Currency | 8 | yes | — | — | — |
| 109 | EarlyEarned | Currency | 8 | yes | — | — | — |
| 110 | TOEarned | Currency | 8 | yes | — | — | — |
| 111 | TOAttBilled | Currency | 8 | yes | — | — | — |
| 112 | CostReimb | Currency | 8 | yes | — | — | — |
| 113 | CBHRev | Currency | 8 | yes | — | — | — |
| 114 | MKRev | Currency | 8 | yes | — | — | — |
| 115 | CBHCom | Currency | 8 | yes | — | — | — |
| 116 | MTRev | Currency | 8 | yes | — | — | — |
| 117 | MTCom | Currency | 8 | yes | — | — | — |
| 118 | KBCom | Currency | 8 | yes | — | — | — |
| 119 | MKCom | Currency | 8 | yes | — | — | — |
| 120 | TOEarnedTr | Boolean | 1 | yes | — | — | — |
| 121 | CostReimbTr | Boolean | 1 | yes | — | — | — |
| 122 | InsertedTrust | Boolean | 1 | yes | — | — | — |
| 123 | TotalHourlyOuts | Currency | 8 | yes | — | — | — |
| 124 | OpenTK | Text | 20 | yes | — | — | allows zero-length string |
| 125 | AdvCostBal | Currency | 8 | yes | — | — | — |
| 126 | AdvFeeBal | Currency | 8 | yes | — | — | — |
| 127 | CostHoldBal | Currency | 8 | yes | — | — | — |
| 128 | BRRev | Currency | 8 | yes | — | — | — |
| 129 | BRCom | Currency | 8 | yes | — | — | — |
| 130 | RLFCom | Currency | 8 | yes | — | — | — |
| 131 | AdvEarned | Currency | 8 | no | — | — | — |
| 132 | RemEarned | Currency | 8 | yes | — | — | — |
| 133 | SumOfCBHRev | Currency | 8 | yes | — | — | — |
| 134 | SumOfMKRev | Currency | 8 | yes | — | — | — |
| 135 | SumOfCBHCom | Currency | 8 | yes | — | — | — |
| 136 | SumOfMTRev | Currency | 8 | yes | — | — | — |
| 137 | SumOfMTCom | Currency | 8 | yes | — | — | — |
| 138 | SumOfKBCom | Currency | 8 | yes | — | — | — |
| 139 | SumOfMKCom | Currency | 8 | yes | — | — | — |
| 140 | SumOfRLFCom | Currency | 8 | yes | — | — | — |
| 141 | SumOfEarlyEarned | Currency | 8 | yes | — | — | — |
| 142 | SumOfTOEarned | Currency | 8 | yes | — | — | — |
| 143 | SumOfTOEarlyAndEarned | Currency | 8 | yes | — | — | — |
| 144 | SumOfCostReimb | Currency | 8 | yes | — | — | — |

**Primary key:** TakeOffID

## Indexes

| Index | Fields | Primary | Unique | Foreign |
|-------|--------|---------|--------|---------|
| UniqueIndex | TakeOffID | yes | yes | no |

## Relationships

_No relationships declared in the database that reference this table._ Check column **lookup** notes above and query joins for implicit foreign keys.

---
*Generated by the extractor's `tableDefs` stage from `schema.json` + `relationships.json`. Structured source of truth: `schema.json`.*
