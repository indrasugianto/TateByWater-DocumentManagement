# Table: Family Law - Divorce *(linked)*

**Linked source:** `ODBC;DRIVER=SQL Server;SERVER=awsql2022dev;APP=Microsoft Office;DATABASE=TateBywater` → `dbo.Family Law - Divorce`

**Row count:** 28

## Columns

| # | Column | Type | Size | Nullable | Default | Key | Notes |
|---|--------|------|------|----------|---------|-----|-------|
| 1 | ID | Long | 4 | no | — | PK | auto-increment |
| 2 | CaseID | Long | 4 | no | — | — | — |
| 3 | C Length at Residence | Text | 255 | yes | — | — | allows zero-length string |
| 4 | C Prior Address | Text | 255 | yes | — | — | allows zero-length string |
| 5 | C Length at Prior Address | Text | 255 | yes | — | — | allows zero-length string |
| 6 | C Length in VA | Text | 255 | yes | — | — | allows zero-length string |
| 7 | C Birthplace | Text | 255 | yes | — | — | allows zero-length string |
| 8 | C Employer | Text | 255 | yes | — | — | allows zero-length string |
| 9 | C Primary Education | Text | 255 | yes | — | — | allows zero-length string |
| 10 | C College | Text | 255 | yes | — | — | allows zero-length string |
| 11 | C Marriage Number | Text | 255 | yes | — | — | allows zero-length string |
| 12 | D Address | Text | 255 | yes | — | — | allows zero-length string |
| 13 | D City | Text | 255 | yes | — | — | allows zero-length string |
| 14 | D State | Text | 255 | yes | — | — | allows zero-length string |
| 15 | D Zip | Text | 255 | yes | — | — | allows zero-length string |
| 16 | D Home Phone | Text | 255 | yes | — | — | allows zero-length string |
| 17 | D Other Phone | Text | 255 | yes | — | — | allows zero-length string |
| 18 | D Email | Text | 255 | yes | — | — | allows zero-length string |
| 19 | D DOB | Text | 255 | yes | — | — | allows zero-length string |
| 20 | D SSN | Text | 255 | yes | — | — | allows zero-length string |
| 21 | D Employer | Text | 255 | yes | — | — | allows zero-length string |
| 22 | D Work Address | Text | 255 | yes | — | — | allows zero-length string |
| 23 | D Work City | Text | 255 | yes | — | — | allows zero-length string |
| 24 | D Work State | Text | 255 | yes | — | — | allows zero-length string |
| 25 | D Work Zip | Text | 255 | yes | — | — | allows zero-length string |
| 26 | D Work Phone | Text | 255 | yes | — | — | allows zero-length string |
| 27 | D Primary Education | Text | 255 | yes | — | — | allows zero-length string |
| 28 | D College | Text | 255 | yes | — | — | allows zero-length string |
| 29 | D Marriage Number | Text | 255 | yes | — | — | allows zero-length string |
| 30 | Date of Marriage | DateTime | 8 | yes | — | — | — |
| 31 | Place of Marriage | Text | 255 | yes | — | — | allows zero-length string |
| 32 | Date of Separation | DateTime | 8 | yes | — | — | — |
| 33 | Length of Separation | Text | 255 | yes | — | — | allows zero-length string |
| 34 | Wife Maiden Name | Text | 255 | yes | — | — | allows zero-length string |
| 35 | Number of Children | Text | 255 | yes | — | — | allows zero-length string |
| 36 | Child Custodian | Text | 255 | yes | — | — | allows zero-length string |
| 37 | C Title | Text | 255 | yes | — | — | allows zero-length string |
| 38 | D Title | Text | 255 | yes | — | — | allows zero-length string |
| 39 | Date of PSA | DateTime | 8 | yes | — | — | — |
| 40 | Place of Last Cohabit | Text | 255 | yes | — | — | allows zero-length string |
| 41 | Divorce Grounds | Text | 255 | yes | — | — | allows zero-length string |
| 42 | FL Court Case No | Text | 255 | yes | — | — | allows zero-length string |
| 43 | Complaint Filed Date | DateTime | 8 | yes | — | — | — |
| 44 | Waiver Date | DateTime | 8 | yes | — | — | — |
| 45 | Publish Dates | DateTime | 8 | yes | — | — | — |
| 46 | Publish Return Date | DateTime | 8 | yes | — | — | — |
| 47 | Complaint Serve Date | DateTime | 8 | yes | — | — | — |
| 48 | Complaint Serve Method | Text | 255 | yes | — | — | allows zero-length string |
| 49 | NOH Serve Date | DateTime | 8 | yes | — | — | — |
| 50 | NOH Serve Method | Text | 255 | yes | — | — | allows zero-length string |
| 51 | Witness | Text | 255 | yes | — | — | allows zero-length string |
| 52 | D_Last_Name | Text | 255 | yes | — | — | allows zero-length string |
| 53 | D_First_Name | Text | 255 | yes | — | — | allows zero-length string |
| 54 | D_BirthPlace | Text | 255 | yes | — | — | allows zero-length string |
| 55 | Uncontested by Affidavit | Boolean | 1 | yes | — | — | — |
| 56 | Waiver of Service | Boolean | 1 | yes | — | — | — |
| 57 | Service by Publication | Boolean | 1 | yes | — | — | — |
| 58 | Sheriff Service | Boolean | 1 | yes | — | — | — |
| 59 | Divorce with MSA | Boolean | 1 | yes | — | — | — |
| 60 | Divorce without MSA | Boolean | 1 | yes | — | — | — |
| 61 | SSMA_TimeStamp | Binary | 8 | no | — | — | — |

**Primary key:** ID

## Indexes

| Index | Fields | Primary | Unique | Foreign |
|-------|--------|---------|--------|---------|
| Family Law - Divorce$PrimaryKey | ID | yes | yes | no |

## Relationships

_No relationships declared in the database that reference this table._ Check column **lookup** notes above and query joins for implicit foreign keys.

---
*Generated by the extractor's `tableDefs` stage from `schema.json` + `relationships.json`. Structured source of truth: `schema.json`.*
