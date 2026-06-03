# Schema Report: TateByWater
---
## Tables
### Bankruptcy (linked)
| Column | Type | Size | Nullable | Default |
|--------|------|------|----------|---------|
| BankruptcyID | Long | 4 | no |  |
| CaseID | Long | 4 | yes |  |
| Chapter | Long | 4 | yes |  |
| Case Filed | DateTime | 8 | yes |  |
| Deadline for Filing Sched | DateTime | 8 | yes |  |
| Document Date for Trustee | DateTime | 8 | yes |  |
| Trustee | Text | 255 | yes |  |
| POCDeadline | DateTime | 8 | yes |  |
| GovtPOC | DateTime | 8 | yes |  |
| Deadline to Object | DateTime | 8 | yes |  |
| BJudge | Text | 255 | yes |  |
| OriginalScheduleDeadline | DateTime | 8 | yes |  |
| PrevBank | Boolean | 1 | yes |  |
| PrevDate | DateTime | 8 | yes |  |
| PrevCaseNumber | Text | 255 | yes |  |
| PrevLocation | Text | 255 | yes |  |
| TrusteeAddress | Text | 255 | yes |  |
| TrusteeCity | Text | 255 | yes |  |
| TrusteeZip | Text | 255 | yes |  |
| TrusteeState | Text | 255 | yes |  |
| TrusteeDocuments | Text | 255 | yes |  |
| ForeTrustee | Text | 255 | yes |  |
| ForeAddress | Text | 255 | yes |  |
| ForeCity | Text | 255 | yes |  |
| ForeState | Text | 255 | yes |  |
| ForeZIP | Text | 255 | yes |  |
| ForePhone | Text | 255 | yes |  |
| ForeFax | Text | 255 | yes |  |
| ForeSaleDate | DateTime | 8 | yes |  |
| ForeTime | DateTime | 8 | yes |  |
| ForeFileNumber | Text | 255 | yes |  |
| TrusteePhone | Text | 255 | yes |  |
| TrusteeFax | Text | 255 | yes |  |
| TrusteeEmail | Text | 255 | yes |  |
| SSMA_TimeStamp | Binary | 8 | no |  |

**Primary Key:** BankruptcyID

**Linked Connect:** `ODBC;DRIVER=SQL Server;SERVER=awsql2022dev;APP=Microsoft Office;DATABASE=TateBywater`

**Indexes:**
- Bankruptcy$PrimaryKey [primary, unique]: BankruptcyID

### Billing (linked)
| Column | Type | Size | Nullable | Default |
|--------|------|------|----------|---------|
| ID | Long | 4 | no |  |
| CaseID | Long | 4 | no |  |
| Balance Due Date | DateTime | 8 | yes |  |
| Past Due | Boolean | 1 | yes |  |
| Long Term Collections | Boolean | 1 | yes |  |
| chkBalanceDue | Boolean | 1 | yes |  |
| Billing Notes | Memo | 0 | yes |  |
| WriteOff | Boolean | 1 | yes |  |
| CostHold | Currency | 8 | yes |  |
| SSMA_TimeStamp | Binary | 8 | no |  |

**Primary Key:** ID

**Linked Connect:** `ODBC;DRIVER=SQL Server;SERVER=awsql2022dev;APP=Microsoft Office;DATABASE=TateBywater`

**Indexes:**
- Billing$PrimaryKey [primary, unique]: ID

### CH13Plans (linked)
| Column | Type | Size | Nullable | Default |
|--------|------|------|----------|---------|
| IDCH13Plans | Long | 4 | no |  |
| IDBankruptcy | Long | 4 | yes |  |
| PlanNr | Long | 4 | yes |  |
| DateFiled | DateTime | 8 | yes |  |
| ConfirmDate | DateTime | 8 | yes |  |
| Notes | Text | 255 | yes |  |
| Confirmed | Boolean | 1 | yes |  |
| Objected | Boolean | 1 | yes |  |
| SSMA_TimeStamp | Binary | 8 | no |  |

**Primary Key:** IDCH13Plans

**Linked Connect:** `ODBC;DRIVER=SQL Server;SERVER=awsql2022dev;APP=Microsoft Office;DATABASE=TateBywater`

**Indexes:**
- CH13Plans$PrimaryKey [primary, unique]: IDCH13Plans

### Disposition (linked)
| Column | Type | Size | Nullable | Default |
|--------|------|------|----------|---------|
| DispoID | Long | 4 | no |  |
| CaseID | Long | 4 | yes |  |
| Disposition | Memo | 0 | yes |  |
| Trial | Boolean | 1 | yes |  |
| Plea | Boolean | 1 | yes |  |
| Not Guilty Dismissed | Boolean | 1 | yes |  |
| Entire np | Boolean | 1 | yes |  |
| PI Settlement Amount | Currency | 8 | yes |  |
| Dispo_Date | DateTime | 8 | yes |  |
| Dispo_Atty | Text | 255 | yes |  |
| DispoJudge | Text | 255 | yes |  |
| DispoOppC | Text | 255 | yes |  |
| Total Earned Fee | Currency | 8 | yes |  |
| SSMA_TimeStamp | Binary | 8 | no |  |

**Primary Key:** DispoID

**Linked Connect:** `ODBC;DRIVER=SQL Server;SERVER=awsql2022dev;APP=Microsoft Office;DATABASE=TateBywater`

**Indexes:**
- Disposition$PrimaryKey [primary, unique]: DispoID

### errMsgs (linked)
| Column | Type | Size | Nullable | Default |
|--------|------|------|----------|---------|
| ID | Long | 4 | no |  |
| Message | Text | 255 | yes |  |
| Buttons | Text | 255 | yes |  |
| Title | Text | 255 | yes |  |

**Primary Key:** ID

**Linked Connect:** `ODBC;DRIVER=SQL Server;SERVER=awsql2022dev;APP=Microsoft Office;DATABASE=TateBywater`

**Indexes:**
- errMsgs$PrimaryKey [primary, unique]: ID

### Family Law - Divorce (linked)
| Column | Type | Size | Nullable | Default |
|--------|------|------|----------|---------|
| ID | Long | 4 | no |  |
| CaseID | Long | 4 | no |  |
| C Length at Residence | Text | 255 | yes |  |
| C Prior Address | Text | 255 | yes |  |
| C Length at Prior Address | Text | 255 | yes |  |
| C Length in VA | Text | 255 | yes |  |
| C Birthplace | Text | 255 | yes |  |
| C Employer | Text | 255 | yes |  |
| C Primary Education | Text | 255 | yes |  |
| C College | Text | 255 | yes |  |
| C Marriage Number | Text | 255 | yes |  |
| D Address | Text | 255 | yes |  |
| D City | Text | 255 | yes |  |
| D State | Text | 255 | yes |  |
| D Zip | Text | 255 | yes |  |
| D Home Phone | Text | 255 | yes |  |
| D Other Phone | Text | 255 | yes |  |
| D Email | Text | 255 | yes |  |
| D DOB | Text | 255 | yes |  |
| D SSN | Text | 255 | yes |  |
| D Employer | Text | 255 | yes |  |
| D Work Address | Text | 255 | yes |  |
| D Work City | Text | 255 | yes |  |
| D Work State | Text | 255 | yes |  |
| D Work Zip | Text | 255 | yes |  |
| D Work Phone | Text | 255 | yes |  |
| D Primary Education | Text | 255 | yes |  |
| D College | Text | 255 | yes |  |
| D Marriage Number | Text | 255 | yes |  |
| Date of Marriage | DateTime | 8 | yes |  |
| Place of Marriage | Text | 255 | yes |  |
| Date of Separation | DateTime | 8 | yes |  |
| Length of Separation | Text | 255 | yes |  |
| Wife Maiden Name | Text | 255 | yes |  |
| Number of Children | Text | 255 | yes |  |
| Child Custodian | Text | 255 | yes |  |
| C Title | Text | 255 | yes |  |
| D Title | Text | 255 | yes |  |
| Date of PSA | DateTime | 8 | yes |  |
| Place of Last Cohabit | Text | 255 | yes |  |
| Divorce Grounds | Text | 255 | yes |  |
| FL Court Case No | Text | 255 | yes |  |
| Complaint Filed Date | DateTime | 8 | yes |  |
| Waiver Date | DateTime | 8 | yes |  |
| Publish Dates | DateTime | 8 | yes |  |
| Publish Return Date | DateTime | 8 | yes |  |
| Complaint Serve Date | DateTime | 8 | yes |  |
| Complaint Serve Method | Text | 255 | yes |  |
| NOH Serve Date | DateTime | 8 | yes |  |
| NOH Serve Method | Text | 255 | yes |  |
| Witness | Text | 255 | yes |  |
| D_Last_Name | Text | 255 | yes |  |
| D_First_Name | Text | 255 | yes |  |
| D_BirthPlace | Text | 255 | yes |  |
| Uncontested by Affidavit | Boolean | 1 | yes |  |
| Waiver of Service | Boolean | 1 | yes |  |
| Service by Publication | Boolean | 1 | yes |  |
| Sheriff Service | Boolean | 1 | yes |  |
| Divorce with MSA | Boolean | 1 | yes |  |
| Divorce without MSA | Boolean | 1 | yes |  |
| SSMA_TimeStamp | Binary | 8 | no |  |

**Primary Key:** ID

**Linked Connect:** `ODBC;DRIVER=SQL Server;SERVER=awsql2022dev;APP=Microsoft Office;DATABASE=TateBywater`

**Indexes:**
- Family Law - Divorce$PrimaryKey [primary, unique]: ID

### Matter and AR (linked)
| Column | Type | Size | Nullable | Default |
|--------|------|------|----------|---------|
| MatterID | Long | 4 | no |  |
| CaseID | Long | 4 | no |  |
| Date2 | DateTime | 8 | yes |  |
| Pay_Outlay | Memo | 0 | yes |  |
| Charge | Currency | 8 | yes |  |
| Payment | Currency | 8 | yes |  |
| FirmPrepaid | Boolean | 1 | yes |  |
| OrderNr | Long | 4 | yes |  |
| InsertPymt | Boolean | 1 | yes |  |
| AdvancedLegal | Boolean | 1 | yes |  |
| SSMA_TimeStamp | Binary | 8 | no |  |
| Creimb | Boolean | 1 | yes |  |

**Primary Key:** MatterID

**Linked Connect:** `ODBC;DRIVER=SQL Server;SERVER=awsql2022dev;APP=Microsoft Office;DATABASE=TateBywater`

**Indexes:**
- Matter and AR$PrimaryKey [primary, unique]: MatterID

### Personal Injury (linked)
| Column | Type | Size | Nullable | Default |
|--------|------|------|----------|---------|
| ID | Long | 4 | no |  |
| CaseID | Long | 4 | no |  |
| ClaimNo1 | Text | 255 | yes |  |
| InsCo1 | Text | 255 | yes |  |
| Adjuster1 | Text | 255 | yes |  |
| Adjuster1 Address | Text | 255 | yes |  |
| Adjuster1 City | Text | 255 | yes |  |
| Adjuster1 State | Text | 255 | yes |  |
| Adjuster1 Zip | Text | 255 | yes |  |
| Adjuster1 Phone | Text | 255 | yes |  |
| Adjuster1 Fax | Text | 255 | yes |  |
| Adjuster1 Email | Text | 255 | yes |  |
| ClaimNo2 | Text | 255 | yes |  |
| InsCo2 | Text | 255 | yes |  |
| Adjuster2 | Text | 255 | yes |  |
| Adjuster2 Address | Text | 255 | yes |  |
| Adjuster2 City | Text | 255 | yes |  |
| Adjuster2 State | Text | 255 | yes |  |
| Adjuster2 Zip | Text | 255 | yes |  |
| Adjuster2 Phone | Text | 255 | yes |  |
| Adjuster2 Fax | Text | 255 | yes |  |
| Adjuster2 Email | Text | 255 | yes |  |
| Filing Date | DateTime | 8 | yes |  |
| Medicare | Boolean | 1 | yes |  |
| Med Pay | Boolean | 1 | yes |  |
| ERISA | Boolean | 1 | yes |  |
| Litigation | Boolean | 1 | yes |  |
| Slip and Fall | Boolean | 1 | yes |  |
| Auto Accident | Boolean | 1 | yes |  |
| Medical Lien | Boolean | 1 | yes |  |
| Assignment | Boolean | 1 | yes |  |
| Med Mal | Boolean | 1 | yes |  |
| DOI | DateTime | 8 | yes |  |
| HealthIns | Text | 255 | yes |  |
| PolicyNo | Text | 255 | yes |  |
| GroupNo | Text | 255 | yes |  |
| csettleper | Text | 255 | yes |  |
| csettlelit | Text | 255 | yes |  |
| location | Text | 255 | yes |  |
| Medicaid | Boolean | 1 | yes |  |
| OppPartyInsured | Text | 255 | yes |  |
| Demand | Currency | 8 | yes |  |
| BriefDescription | Memo | 0 | yes |  |
| PIState | Text | 255 | yes |  |
| AutoCarrier | Text | 255 | yes |  |
| AutoPolicyNo | Text | 255 | yes |  |
| UnderinsLimits | Currency | 8 | yes |  |
| MaxMed | Currency | 8 | yes |  |
| PISOL | DateTime | 8 | yes |  |
| PolicyNo1 | Text | 255 | yes |  |
| AdjusterExt | Text | 255 | yes |  |
| CompltServed | Boolean | 1 | yes |  |
| ServedDate | DateTime | 8 | yes |  |
| OtherDriver | Text | 255 | yes |  |
| SSMA_TimeStamp | Binary | 8 | no |  |
| PIStatus | Text | 50 | yes |  |
| LiabilityLimit | Currency | 8 | yes |  |

**Primary Key:** ID

**Linked Connect:** `ODBC;DRIVER=SQL Server;SERVER=awsql2022dev;APP=Microsoft Office;DATABASE=TateBywater`

**Indexes:**
- Personal Injury$PrimaryKey [primary, unique]: ID

### ProofOfClaims (linked)
| Column | Type | Size | Nullable | Default |
|--------|------|------|----------|---------|
| IDProofOfClaims | Long | 4 | no |  |
| IDBankruptcy | Long | 4 | yes |  |
| ClaimNr | Long | 4 | yes |  |
| DateFiled | DateTime | 8 | yes |  |
| CreditorName | Text | 255 | yes |  |
| Secured | Currency | 8 | yes |  |
| Priority | Currency | 8 | yes |  |
| Unsecured | Currency | 8 | yes |  |
| Arrears | Currency | 8 | yes |  |

**Primary Key:** IDProofOfClaims

**Linked Connect:** `ODBC;DRIVER=SQL Server;SERVER=awsql2022dev;APP=Microsoft Office;DATABASE=TateBywater`

**Indexes:**
- ProofOfClaims$PrimaryKey [primary, unique]: IDProofOfClaims

### TB Intakes (linked)
| Column | Type | Size | Nullable | Default |
|--------|------|------|----------|---------|
| ID | Long | 4 | no |  |
| GI Last Name | Text | 255 | yes |  |
| GI First Name | Text | 255 | yes |  |
| GI phone | Text | 255 | yes |  |
| GI Date | DateTime | 8 | yes |  |
| GI Practice Area | Text | 255 | yes |  |
| GI Individual Referrer | Text | 255 | yes |  |
| GI Comments | Memo | 0 | yes |  |
| GI No Further Action | Boolean | 1 | yes |  |
| GI Open | Boolean | 1 | yes |  |
| GI Open Date | DateTime | 8 | yes |  |
| GI Referral | Long | 4 | yes |  |
| ReasonDintHire | Memo | 0 | yes |  |
| FollowUpDate | DateTime | 8 | yes |  |
| Attorny | Text | 255 | yes |  |
| QuotedFee | Currency | 8 | yes |  |
| Scan Location GI | Memo | 0 | yes |  |
| Scanned GI | Boolean | 1 | yes |  |
| GIDOB | DateTime | 8 | yes |  |
| GI Matter | Text | 255 | yes |  |
| SSMA_TimeStamp | Binary | 8 | no |  |

**Primary Key:** ID

**Linked Connect:** `ODBC;DRIVER=SQL Server;SERVER=awsql2022dev;APP=Microsoft Office;DATABASE=TateBywater`

**Indexes:**
- TB Intakes$PrimaryKey [primary, unique]: ID

### TB Time Keeping (linked)
| Column | Type | Size | Nullable | Default |
|--------|------|------|----------|---------|
| Bill_ID | Long | 4 | no |  |
| CaseID | Long | 4 | yes |  |
| Bill Sent | DateTime | 8 | yes |  |
| Bill Paid | DateTime | 8 | yes |  |
| Bill Closed | Boolean | 1 | yes |  |
| BilL Closed Date | DateTime | 8 | yes |  |
| Discount | Currency | 8 | yes |  |
| Bill Number | Text | 255 | yes |  |
| IANumber | Text | 255 | yes |  |
| Bill Open | DateTime | 8 | yes |  |
| TimeNotes | Memo | 0 | yes |  |
| TKLocked | Boolean | 1 | yes |  |
| InvoiceTotalAdvance | Boolean | 1 | yes |  |
| InvoiceExceedsTrust | Boolean | 1 | yes |  |
| StatementLessTrust | Boolean | 1 | yes |  |
| TrustatClose | Currency | 8 | yes |  |
| InvoiceNoAdvance | Boolean | 1 | yes |  |
| ARatClose | Currency | 8 | yes |  |
| AdvBalanceatClose | Currency | 8 | yes |  |
| ReplenishBalanceatClose | Currency | 8 | yes |  |
| AdvCostBal | Currency | 8 | yes |  |
| AdvFeesBal | Currency | 8 | yes |  |
| OutsAdvDue | Currency | 8 | yes |  |
| CostResBalance | Currency | 8 | yes |  |
| InvoiceCostHold | Boolean | 1 | yes |  |
| InvoiceAdvCostFee | Boolean | 1 | yes |  |
| SSMA_TimeStamp | Binary | 8 | no |  |

**Primary Key:** Bill_ID

**Linked Connect:** `ODBC;DRIVER=SQL Server;SERVER=awsql2022dev;APP=Microsoft Office;DATABASE=TateBywater`

**Indexes:**
- TB Time Keeping$PrimaryKey [primary, unique]: Bill_ID

### tbl_CtCaseNumbers (linked)
| Column | Type | Size | Nullable | Default |
|--------|------|------|----------|---------|
| CtCaseNoID | Long | 4 | no |  |
| CaseID | Long | 4 | yes |  |
| Matter_Charge | Text | 255 | yes |  |
| CtNumber | Text | 255 | yes |  |
| District | Boolean | 1 | yes |  |
| Circuit | Boolean | 1 | yes |  |
| CodeSection | Text | 255 | yes |  |
| SSMA_TimeStamp | Binary | 8 | no |  |

**Primary Key:** CtCaseNoID

**Linked Connect:** `ODBC;DRIVER=SQL Server;SERVER=awsql2022dev;APP=Microsoft Office;DATABASE=TateBywater`

**Indexes:**
- tbl_CtCaseNumbers$PrimaryKey [primary, unique]: CtCaseNoID

### tbl_InvoiceSent (linked)
| Column | Type | Size | Nullable | Default |
|--------|------|------|----------|---------|
| InvoiceSentID | Long | 4 | no |  |
| CaseID | Long | 4 | yes |  |
| InvSent | DateTime | 8 | yes |  |
| InvoiceNumber | Text | 255 | yes |  |
| TK Sent | Boolean | 1 | yes |  |
| TKDate | DateTime | 8 | yes |  |
| InvSentNotes | Text | 255 | yes |  |
| InvBalance | Currency | 8 | yes |  |
| TKNumber | Text | 255 | yes |  |
| TKBalance | Currency | 8 | yes |  |
| ClientCall | DateTime | 8 | yes |  |
| SSMA_TimeStamp | Binary | 8 | no |  |

**Primary Key:** InvoiceSentID

**Linked Connect:** `ODBC;DRIVER=SQL Server;SERVER=awsql2022dev;APP=Microsoft Office;DATABASE=TateBywater`

**Indexes:**
- tbl_InvoiceSent$PrimaryKey [primary, unique]: InvoiceSentID

### tblAccessType (linked)
| Column | Type | Size | Nullable | Default |
|--------|------|------|----------|---------|
| ID | Long | 4 | no |  |
| AccessType | Long | 4 | yes |  |
| AccessDescription | Text | 255 | yes |  |
| AdminPane | Boolean | 1 | yes |  |
| SSMA_TimeStamp | Binary | 8 | no |  |

**Primary Key:** ID

**Linked Connect:** `ODBC;DRIVER=SQL Server;SERVER=awsql2022dev;APP=Microsoft Office;DATABASE=TateBywater`

**Indexes:**
- tblAccessType$PrimaryKey [primary, unique]: ID

### TblActionNeeded (linked)
| Column | Type | Size | Nullable | Default |
|--------|------|------|----------|---------|
| ActionNeededID | Long | 4 | no |  |
| CaseID | Long | 4 | yes |  |
| ActionNeededDet | Text | 255 | yes |  |
| ActionComp | Boolean | 1 | yes |  |
| SSMA_TimeStamp | Binary | 8 | no |  |
| DateComp | Text | 10 | yes |  |
| DateComp1 | DateTime | 8 | yes |  |
| ActPerson | Text | 10 | yes |  |
| StartDate | DateTime | 8 | yes |  |

**Primary Key:** ActionNeededID

**Linked Connect:** `ODBC;DRIVER=SQL Server;SERVER=awsql2022dev;APP=Microsoft Office;DATABASE=TateBywater`

**Indexes:**
- TblActionNeeded$PrimaryKey [primary, unique]: ActionNeededID

### tblAttorneys (linked)
| Column | Type | Size | Nullable | Default |
|--------|------|------|----------|---------|
| AttysID | Long | 4 | no |  |
| AttyName | Text | 255 | yes |  |
| AttyEmail | Text | 255 | yes |  |
| VABar | Text | 255 | yes |  |
| MDBar | Text | 255 | yes |  |
| DCBar | Text | 255 | yes |  |
| USBankBar | Text | 255 | yes |  |
| AttyInitials | Text | 255 | yes |  |
| AttyCel | Text | 255 | yes |  |

**Primary Key:** AttysID

**Linked Connect:** `ODBC;DRIVER=SQL Server;SERVER=awsql2022dev;APP=Microsoft Office;DATABASE=TateBywater`

**Indexes:**
- tblAttorneys$PrimaryKey [primary, unique]: AttysID

### tblCalls (linked)
| Column | Type | Size | Nullable | Default |
|--------|------|------|----------|---------|
| CallID | Long | 4 | no |  |
| CFirstName | Text | 255 | yes |  |
| CLastName | Text | 255 | yes |  |
| CDate | DateTime | 8 | yes |  |
| CPracticeArea | Text | 255 | yes |  |
| CReferral | Text | 255 | yes |  |
| Consult | Boolean | 1 | yes |  |
| CAtty | Text | 255 | yes |  |
| CallTime | DateTime | 8 | yes |  |
| HiredFirm | Boolean | 1 | yes |  |
| CallMatter | Text | 255 | yes |  |
| CallComments | Memo | 0 | yes |  |
| Transferred to Atty | Text | 255 | yes |  |
| CPhone | Text | 255 | yes |  |
| SchedDate | DateTime | 8 | yes |  |
| SchedTime | DateTime | 8 | yes |  |
| ClientType | Text | 255 | yes |  |
| attyEmail | Text | 255 | yes |  |
| SSMA_TimeStamp | Binary | 8 | no |  |
| CPhoneType | Text | 255 | yes |  |
| CPhoneExt | Text | 255 | yes |  |

**Primary Key:** CallID

**Linked Connect:** `ODBC;DRIVER=SQL Server;SERVER=awsql2022dev;APP=Microsoft Office;DATABASE=TateBywater`

**Indexes:**
- tblCalls$PrimaryKey [primary, unique]: CallID

### tblCase (linked)
| Column | Type | Size | Nullable | Default |
|--------|------|------|----------|---------|
| CaseID | Long | 4 | no |  |
| Last_Name | Text | 255 | yes |  |
| First_Name | Text | 255 | yes |  |
| CaseOpenDate | DateTime | 8 | yes |  |
| Closed | Boolean | 1 | yes |  |
| Clsdate | DateTime | 8 | yes |  |
| Extended_Ledger | Text | 255 | yes |  |
| Case_Letter | Text | 255 | yes |  |
| yr | Text | 254 | yes |  |
| Number_ | Long | 4 | yes |  |
| Orig_Atty | Text | 255 | yes |  |
| Address | Text | 255 | yes |  |
| CourtCaseNo | Text | 255 | yes |  |
| City | Text | 255 | yes |  |
| FamilyLaw | Boolean | 1 | yes |  |
| State | Text | 255 | yes |  |
| Zip | Text | 255 | yes |  |
| Country | Text | 255 | yes |  |
| HmPhone | Text | 255 | yes |  |
| Action | Text | 255 | yes |  |
| OtherPhone | Text | 255 | yes |  |
| Fax | Text | 255 | yes |  |
| WkPhone | Text | 255 | yes |  |
| Comments | Memo | 0 | yes |  |
| Email | Text | 255 | yes |  |
| Referral | Text | 255 | yes |  |
| Individual Referrer | Text | 255 | yes |  |
| Retainer | Currency | 8 | yes |  |
| Matter_type | Text | 255 | yes |  |
| SOL | DateTime | 8 | yes |  |
| Court | Text | 255 | yes |  |
| CType | Text | 255 | yes |  |
| POfc | Text | 255 | yes |  |
| ComplainingWitness | Text | 255 | yes |  |
| DOB | DateTime | 8 | yes |  |
| WkAddress | Text | 255 | yes |  |
| WkCity | Text | 255 | yes |  |
| WkState | Text | 255 | yes |  |
| WkZip | Text | 255 | yes |  |
| Pro Bono | Boolean | 1 | yes |  |
| HandlingAtty_Case | Text | 255 | yes |  |
| Action_Needed_on_Payment | Boolean | 1 | yes |  |
| SSN | Text | 255 | yes |  |
| Employer Name | Text | 255 | yes |  |
| Last Updated Contact Info | DateTime | 8 | yes |  |
| Ocounsel | Text | 255 | yes |  |
| Firm | Text | 255 | yes |  |
| OC Address | Text | 255 | yes |  |
| OC City | Text | 255 | yes |  |
| OC State | Text | 255 | yes |  |
| OC Zip | Text | 255 | yes |  |
| OC Phone | Text | 255 | yes |  |
| OC Email | Text | 255 | yes |  |
| OC Fax | Text | 255 | yes |  |
| Pro Bono PM | Text | 255 | yes |  |
| Pro Bono JRT | Text | 255 | yes |  |
| ContingencyFee | Boolean | 1 | yes |  |
| AuthorityToTalkTo | Memo | 0 | yes |  |
| Hourly | Boolean | 1 | yes |  |
| Contingency | Boolean | 1 | yes |  |
| Hybrid | Boolean | 1 | yes |  |
| Family-Law | Boolean | 1 | yes |  |
| Fixed | Boolean | 1 | yes |  |
| Scan | Boolean | 1 | yes |  |
| Scan Location | Memo | 0 | yes |  |
| ScanNotAvail | Boolean | 1 | yes |  |
| ParaLegal | Text | 255 | yes |  |
| Spanish | Boolean | 1 | yes |  |
| Offdate | DateTime | 8 | yes |  |
| CostHold | Currency | 8 | yes |  |
| CltNarrative | Memo | 0 | yes |  |
| ARTrustZero | Boolean | 1 | yes |  |
| F73 | Text | 255 | yes |  |
| F74 | Text | 255 | yes |  |
| F75 | Text | 255 | yes |  |
| F76 | Text | 255 | yes |  |
| PhName1 | Text | 255 | yes |  |
| PhName2 | Text | 255 | yes |  |
| LengthRes | Text | 255 | yes |  |
| LengthEmp | Text | 255 | yes |  |
| LegalStatus | Text | 255 | yes |  |
| CurrentBond | Text | 255 | yes |  |
| CrRecord | Memo | 0 | yes |  |
| TrustChronMemo | Memo | 0 | yes |  |
| Executor | Text | 255 | yes |  |
| RetainerReimb | Boolean | 1 | yes |  |
| RetReimbAmount | Currency | 8 | yes |  |
| Reviewable | Boolean | 1 | yes |  |
| ReviewReq | DateTime | 8 | yes |  |
| ReviewReceivedDate | DateTime | 8 | yes |  |
| ReviewReceived | Boolean | 1 | yes |  |
| Testimonial | Memo | 0 | yes |  |
| ReviewFollowUp | DateTime | 8 | yes |  |
| Stars | Long | 4 | yes |  |
| Review Source | Text | 255 | yes |  |
| Review Date | DateTime | 8 | yes |  |
| Title | Text | 255 | yes |  |
| OPartyLast | Text | 255 | yes |  |
| OPartyFirst | Text | 255 | yes |  |
| OPartyDOB | DateTime | 8 | yes |  |
| SSMA_TimeStamp | Binary | 8 | no |  |
| PartnerRate | Currency | 8 | yes |  |
| AssocRate | Currency | 8 | yes |  |

**Primary Key:** CaseID

**Linked Connect:** `ODBC;DRIVER=SQL Server;SERVER=awsql2022dev;APP=Microsoft Office;DATABASE=TateBywater`

**Indexes:**
- tblCase$PrimaryKey [primary, unique]: CaseID

### tblCaseDocuments (linked)
| Column | Type | Size | Nullable | Default |
|--------|------|------|----------|---------|
| CaseDocumentID | Long | 4 | no |  |
| CaseID | Long | 4 | no |  |
| DocumentType | Text | 250 | no |  |
| DocumentFileName | Memo | 0 | no |  |
| CreatedOn | DateTime | 8 | no |  |

**Primary Key:** CaseDocumentID

**Linked Connect:** `ODBC;DRIVER=SQL Server;SERVER=awsql2022dev;APP=Microsoft Office;DATABASE=TateBywater`

**Indexes:**
- PK_tblCaseDocuments [primary, unique]: CaseDocumentID

### tblChild (linked)
| Column | Type | Size | Nullable | Default |
|--------|------|------|----------|---------|
| Child_ID | Long | 4 | no |  |
| FamilyLaw_ID | Long | 4 | yes |  |
| ChildName | Text | 255 | yes |  |
| DOB_child | DateTime | 8 | yes |  |

**Primary Key:** Child_ID

**Linked Connect:** `ODBC;DRIVER=SQL Server;SERVER=awsql2022dev;APP=Microsoft Office;DATABASE=TateBywater`

**Indexes:**
- tblChild$PrimaryKey [primary, unique]: Child_ID

### tblDocumentRootDirectory (linked)
| Column | Type | Size | Nullable | Default |
|--------|------|------|----------|---------|
| DocumentRootDirectoryID | Long | 4 | no |  |
| DocumentRootDirectory | Memo | 0 | no |  |
| ScannerDirectory | Memo | 0 | no |  |
| DocumentRootNaming | Memo | 0 | no |  |
| DocumentClosedNaming | Memo | 0 | no |  |
| AllInvoicesDirectory | Memo | 0 | no |  |
| AllInvoicesNaming | Memo | 0 | no |  |
| ClosedFileScanDirectory | Memo | 0 | no |  |
| ClosedFileScanNaming | Memo | 0 | no |  |
| IntakeDirectory | Memo | 0 | no |  |

**Primary Key:** DocumentRootDirectoryID

**Linked Connect:** `ODBC;DRIVER=SQL Server;SERVER=awsql2022dev;APP=Microsoft Office;DATABASE=TateBywater`

**Indexes:**
- PK_tblDocumentRootDirectory [primary, unique]: DocumentRootDirectoryID

### tblDocumentTypes (linked)
| Column | Type | Size | Nullable | Default |
|--------|------|------|----------|---------|
| DocumentTypeID | Long | 4 | no |  |
| DocumentType | Text | 250 | no |  |
| DocumentNamingRule | Memo | 0 | no |  |
| DocumentFolder | Memo | 0 | no |  |
| SortOrder | Long | 4 | no |  |
| IsVisible | Boolean | 1 | no |  |

**Primary Key:** DocumentTypeID

**Linked Connect:** `ODBC;DRIVER=SQL Server;SERVER=awsql2022dev;APP=Microsoft Office;DATABASE=TateBywater`

**Indexes:**
- IX_tblDocumentTypes_DocumentType [unique]: DocumentType
- PK_tblDocumentTypes [primary, unique]: DocumentTypeID

### tblDropD (linked)
| Column | Type | Size | Nullable | Default |
|--------|------|------|----------|---------|
| DropID | Long | 4 | no |  |
| Code | Text | 255 | yes |  |
| CodeVal | Text | 255 | yes |  |
| FieldName | Text | 255 | yes |  |
| Description | Text | 255 | yes |  |
| SortOrder | Long | 4 | yes |  |
| Description1 | Text | 255 | yes |  |
| SortOrder1 | Long | 4 | yes |  |

**Primary Key:** DropID

**Linked Connect:** `ODBC;DRIVER=SQL Server;SERVER=awsql2022dev;APP=Microsoft Office;DATABASE=TateBywater`

**Indexes:**
- tblDropD$PrimaryKey [primary, unique]: DropID

### tblFields (linked)
| Column | Type | Size | Nullable | Default |
|--------|------|------|----------|---------|
| Object | Text | 55 | yes |  |
| FieldName | Text | 55 | yes |  |
| FieldType | Text | 20 | yes |  |
| FieldSize | Long | 4 | yes |  |
| FieldAttributes | Long | 4 | yes |  |
| FldDescription | Text | 20 | yes |  |

**Linked Connect:** `ODBC;DRIVER=SQL Server;SERVER=awsql2022dev;APP=Microsoft Office;DATABASE=TateBywater`

### tblFormAccessMapping (linked)
| Column | Type | Size | Nullable | Default |
|--------|------|------|----------|---------|
| ID | Long | 4 | no |  |
| FormName | Text | 255 | yes |  |
| MinimumAccess_Show | Long | 4 | yes |  |
| MinimumAccess_Edit | Long | 4 | yes |  |

**Primary Key:** ID

**Linked Connect:** `ODBC;DRIVER=SQL Server;SERVER=awsql2022dev;APP=Microsoft Office;DATABASE=TateBywater`

**Indexes:**
- tblFormAccessMapping$PrimaryKey [primary, unique]: ID

### tblHearingDate (linked)
| Column | Type | Size | Nullable | Default |
|--------|------|------|----------|---------|
| HearingID | Long | 4 | no |  |
| CaseID | Long | 4 | yes |  |
| Hearing_Date | DateTime | 8 | yes |  |
| HearingType | Text | 255 | yes |  |
| HearingTime | DateTime | 8 | yes |  |
| HrgResult | Text | 255 | yes |  |
| HrgCal | Boolean | 1 | yes |  |
| Verified | Boolean | 1 | yes |  |
| ClientPresent | Boolean | 1 | yes |  |
| Reminder | Text | 255 | yes |  |
| ReminderCheck | Boolean | 1 | yes |  |
| SSMA_TimeStamp | Binary | 8 | no |  |

**Primary Key:** HearingID

**Linked Connect:** `ODBC;DRIVER=SQL Server;SERVER=awsql2022dev;APP=Microsoft Office;DATABASE=TateBywater`

**Indexes:**
- tblHearingDate$PrimaryKey [primary, unique]: HearingID

### Tblmsgbox (linked)
| Column | Type | Size | Nullable | Default |
|--------|------|------|----------|---------|
| ID | Long | 4 | no |  |
| AccessName | Text | 255 | yes |  |
| MessageName | Text | 255 | yes |  |

**Linked Connect:** `ODBC;DRIVER=SQL Server;SERVER=awsql2022dev;APP=Microsoft Office;DATABASE=TateBywater`

### tblNotes (linked)
| Column | Type | Size | Nullable | Default |
|--------|------|------|----------|---------|
| IDNotes | Long | 4 | no |  |
| CaseID | Long | 4 | yes |  |
| NoteDate | DateTime | 8 | yes |  |
| NotePerson | Text | 255 | yes |  |
| NoteDescription | Memo | 0 | yes |  |
| NoteTime | Long | 4 | yes |  |
| SSMA_TimeStamp | Binary | 8 | no |  |

**Primary Key:** IDNotes

**Linked Connect:** `ODBC;DRIVER=SQL Server;SERVER=awsql2022dev;APP=Microsoft Office;DATABASE=TateBywater`

**Indexes:**
- tblNotes$PrimaryKey [primary, unique]: IDNotes

### tblPersInjDemand (linked)
| Column | Type | Size | Nullable | Default |
|--------|------|------|----------|---------|
| PIDemandID | Long | 4 | no |  |
| ID | Long | 4 | yes |  |
| DemandDate | DateTime | 8 | yes |  |
| DemandParty | Text | 255 | yes |  |
| DemandAmount | Currency | 8 | yes |  |

**Primary Key:** PIDemandID

**Linked Connect:** `ODBC;DRIVER=SQL Server;SERVER=awsql2022dev;APP=Microsoft Office;DATABASE=TateBywater`

**Indexes:**
- tblPersInjDemand$PrimaryKey [primary, unique]: PIDemandID

### tblPersInjLog (linked)
| Column | Type | Size | Nullable | Default |
|--------|------|------|----------|---------|
| PersInjLogID | Long | 4 | no |  |
| ID | Long | 4 | yes |  |
| EventDate | DateTime | 8 | yes |  |
| EventDescription | Memo | 0 | yes |  |
| SSMA_TimeStamp | Binary | 8 | no |  |

**Primary Key:** PersInjLogID

**Linked Connect:** `ODBC;DRIVER=SQL Server;SERVER=awsql2022dev;APP=Microsoft Office;DATABASE=TateBywater`

**Indexes:**
- tblPersInjLog$PrimaryKey [primary, unique]: PersInjLogID

### tblPersInjProv (linked)
| Column | Type | Size | Nullable | Default |
|--------|------|------|----------|---------|
| PIProviderID | Long | 4 | no |  |
| ID | Long | 4 | yes |  |
| Provider | Text | 255 | yes |  |
| ReqDate | DateTime | 8 | yes |  |
| RcvDate | DateTime | 8 | yes |  |
| PBillAmount | Currency | 8 | yes |  |
| Lien | Boolean | 1 | yes |  |
| SSMA_TimeStamp | Binary | 8 | no |  |
| PBillRed | Currency | 8 | yes |  |

**Primary Key:** PIProviderID

**Linked Connect:** `ODBC;DRIVER=SQL Server;SERVER=awsql2022dev;APP=Microsoft Office;DATABASE=TateBywater`

**Indexes:**
- tblPersInjProv$PrimaryKey [primary, unique]: PIProviderID

### tblPrevBank (linked)
| Column | Type | Size | Nullable | Default |
|--------|------|------|----------|---------|
| IDPrevBank | Long | 4 | no |  |
| IDBankruptcy | Long | 4 | yes |  |
| PrevDate | DateTime | 8 | yes |  |
| PrevCaseNumber | Text | 255 | yes |  |
| PrevLocation | Text | 255 | yes |  |
| PChapter | Text | 255 | yes |  |

**Primary Key:** IDPrevBank

**Linked Connect:** `ODBC;DRIVER=SQL Server;SERVER=awsql2022dev;APP=Microsoft Office;DATABASE=TateBywater`

**Indexes:**
- tblPrevBank$PrimaryKey [primary, unique]: IDPrevBank

### tblReceipts (linked)
| Column | Type | Size | Nullable | Default |
|--------|------|------|----------|---------|
| ReceiptID | Long | 4 | no |  |
| RDate | DateTime | 8 | yes |  |
| RFrom | Text | 255 | yes |  |
| RFor | Text | 255 | yes |  |
| RAmount | Currency | 8 | yes |  |
| RMatter | Text | 255 | yes |  |
| RDue | Currency | 8 | yes |  |
| RCash | Boolean | 1 | yes |  |
| RCC | Boolean | 1 | yes |  |
| RCheck | Boolean | 1 | yes |  |
| RLock | Boolean | 1 | yes |  |
| RCheckNumber | Text | 255 | yes |  |
| SSMA_TimeStamp | Binary | 8 | no |  |

**Primary Key:** ReceiptID

**Linked Connect:** `ODBC;DRIVER=SQL Server;SERVER=awsql2022dev;APP=Microsoft Office;DATABASE=TateBywater`

**Indexes:**
- tblReceipts$PrimaryKey [primary, unique]: ReceiptID

### tblScans (linked)
| Column | Type | Size | Nullable | Default |
|--------|------|------|----------|---------|
| ScansID | Long | 4 | no |  |
| CaseID | Long | 4 | yes |  |
| ScanLocation | Memo | 0 | yes |  |
| TypeofScan | Text | 255 | yes |  |
| SSMA_TimeStamp | Binary | 8 | no |  |

**Primary Key:** ScansID

**Linked Connect:** `ODBC;DRIVER=SQL Server;SERVER=awsql2022dev;APP=Microsoft Office;DATABASE=TateBywater`

**Indexes:**
- tblScans$PrimaryKey [primary, unique]: ScansID

### tblTakeOff (linked)
| Column | Type | Size | Nullable | Default |
|--------|------|------|----------|---------|
| TakeOffID | Long | 4 | no |  |
| CaseID | Long | 4 | yes |  |
| TakeOffMonthID | Long | 4 | no |  |
| AvailBalance | Currency | 8 | yes |  |
| TotalUnCashedChks | Currency | 8 | yes |  |
| TotalUnclearedDeps | Currency | 8 | yes |  |
| TotalAdvancedAR | Currency | 8 | yes |  |
| EarlyEarned | Currency | 8 | yes |  |
| TOEarned | Currency | 8 | yes |  |
| CostReimb | Currency | 8 | yes |  |
| CBHRev | Currency | 8 | yes |  |
| MKRev | Currency | 8 | yes |  |
| CBHCom | Currency | 8 | yes |  |
| MTRev | Currency | 8 | yes |  |
| MTCom | Currency | 8 | yes |  |
| KBCom | Currency | 8 | yes |  |
| MKCom | Currency | 8 | yes |  |
| TOEarnedTr | Boolean | 1 | yes |  |
| CostReimbTr | Boolean | 1 | yes |  |
| InsertedTrust | Boolean | 1 | yes |  |
| TotalHourlyOuts | Currency | 8 | yes |  |
| OpenTK | Text | 20 | yes |  |
| AdvCostBal | Currency | 8 | yes |  |
| AdvFeeBal | Currency | 8 | yes |  |
| CostHoldBal | Currency | 8 | yes |  |
| BRRev | Currency | 8 | yes |  |
| BRCom | Currency | 8 | yes |  |
| RLFCom | Currency | 8 | yes |  |
| SSMA_TimeStamp | Binary | 8 | no |  |
| TOBilled | Currency | 8 | yes |  |
| TOAttBilled | Currency | 8 | yes |  |

**Primary Key:** TakeOffID

**Linked Connect:** `ODBC;DRIVER=SQL Server;SERVER=awsql2022dev;APP=Microsoft Office;DATABASE=TateBywater`

**Indexes:**
- tblTakeOff$PrimaryKey [primary, unique]: TakeOffID

### tblTakeOffMonth (linked)
| Column | Type | Size | Nullable | Default |
|--------|------|------|----------|---------|
| TakeOffMonthID | Long | 4 | no |  |
| TakeOffDate | DateTime | 8 | yes |  |
| WF Balance | Currency | 8 | yes |  |
| SumUncashed | Currency | 8 | yes |  |
| SumUncleared | Currency | 8 | yes |  |
| WFplusuncashed | Currency | 8 | yes |  |
| WFActual | Currency | 8 | yes |  |
| ReconcileValue | Currency | 8 | yes |  |
| DaleBalance | Currency | 8 | yes |  |
| DaleActual | Currency | 8 | yes |  |
| SomBalance | Currency | 8 | yes |  |
| SomActual | Currency | 8 | yes |  |
| CombinedTrust | Currency | 8 | yes |  |
| AccReconciled | Boolean | 1 | yes |  |
| JRTFees | Currency | 8 | yes |  |
| DEBFees | Currency | 8 | yes |  |
| GBFFees | Currency | 8 | yes |  |
| PMFees | Currency | 8 | yes |  |
| TDTFees | Currency | 8 | yes |  |
| CBHFees | Currency | 8 | yes |  |
| MKFees | Currency | 8 | yes |  |
| MTFees | Currency | 8 | yes |  |
| KDBFees | Currency | 8 | yes |  |
| TotalTOEarned | Currency | 8 | yes |  |
| TotalTOCostReimb | Currency | 8 | yes |  |
| TotalTOCommissions | Currency | 8 | yes |  |
| TotalCBHCommissions | Text | 255 | yes |  |
| TotalCBHRev | Currency | 8 | yes |  |
| TotalMKCommissions | Currency | 8 | yes |  |
| TotalMKRev | Currency | 8 | yes |  |
| TotalMTCommissions | Currency | 8 | yes |  |
| TotalKDBCommissions | Currency | 8 | yes |  |
| FeeDataInserted | Boolean | 1 | yes |  |
| LastWF | Text | 255 | yes |  |
| NHFees | Currency | 8 | yes |  |
| RLFFees | Currency | 8 | yes |  |
| RLFDeedFees | Currency | 8 | yes |  |
| TotalRLFCommissions | Currency | 8 | yes |  |
| SSMA_TimeStamp | Binary | 8 | no |  |
| JFFees | Currency | 8 | yes |  |
| WNEFees | Currency | 8 | yes |  |

**Primary Key:** TakeOffMonthID

**Linked Connect:** `ODBC;DRIVER=SQL Server;SERVER=awsql2022dev;APP=Microsoft Office;DATABASE=TateBywater`

**Indexes:**
- tblTakeOffMonth$PrimaryKey [primary, unique]: TakeOffMonthID

### tblTimeTableDetail (linked)
| Column | Type | Size | Nullable | Default |
|--------|------|------|----------|---------|
| Time_ID | Long | 4 | no |  |
| Bill_ID | Long | 4 | yes |  |
| Tdate | DateTime | 8 | yes |  |
| Description | Memo | 0 | yes |  |
| Tatty | Text | 255 | yes |  |
| Rate | Currency | 8 | yes |  |
| Time_ | Double | 8 | yes |  |
| SSMA_TimeStamp | Binary | 8 | no |  |

**Primary Key:** Time_ID

**Linked Connect:** `ODBC;DRIVER=SQL Server;SERVER=awsql2022dev;APP=Microsoft Office;DATABASE=TateBywater`

**Indexes:**
- tblTimeTableDetail$PrimaryKey [primary, unique]: Time_ID

### tblUsers (linked)
| Column | Type | Size | Nullable | Default |
|--------|------|------|----------|---------|
| ID | Long | 4 | no |  |
| UserID | Text | 255 | no |  |
| PWD | Text | 255 | no |  |
| Access | Long | 4 | yes |  |

**Primary Key:** ID

**Linked Connect:** `ODBC;DRIVER=SQL Server;SERVER=awsql2022dev;APP=Microsoft Office;DATABASE=TateBywater`

**Indexes:**
- tblUsers$PrimaryKey [primary, unique]: ID

### tblYearMap (linked)
| Column | Type | Size | Nullable | Default |
|--------|------|------|----------|---------|
| ID | Long | 4 | no |  |
| YearNum | Double | 8 | yes |  |
| YearID | Text | 255 | yes |  |
| SSMA_TimeStamp | Binary | 8 | no |  |

**Linked Connect:** `ODBC;DRIVER=SQL Server;SERVER=awsql2022dev;APP=Microsoft Office;DATABASE=TateBywater`

### Trust Account (linked)
| Column | Type | Size | Nullable | Default |
|--------|------|------|----------|---------|
| TrustAccountID | Long | 4 | no |  |
| CaseID | Long | 4 | yes |  |
| TDate | DateTime | 8 | yes |  |
| TMatter | Memo | 0 | yes |  |
| Debit | Currency | 8 | yes |  |
| Credit | Currency | 8 | yes |  |
| CheckCashed | Boolean | 1 | yes |  |
| CheckNumber | Text | 255 | yes |  |
| DepCleared | Boolean | 1 | yes |  |
| Reconciled | Boolean | 1 | yes |  |
| OrderNr | Long | 4 | yes |  |
| CostReimb | Boolean | 1 | yes |  |
| AdvFee | Boolean | 1 | yes |  |
| SSMA_TimeStamp | Binary | 8 | no |  |

**Primary Key:** TrustAccountID

**Linked Connect:** `ODBC;DRIVER=SQL Server;SERVER=awsql2022dev;APP=Microsoft Office;DATABASE=TateBywater`

**Indexes:**
- Trust Account$PrimaryKey [primary, unique]: TrustAccountID

### vw_advanced_payments (linked)
| Column | Type | Size | Nullable | Default |
|--------|------|------|----------|---------|
| Name | Memo | 0 | no |  |
| FileNumber | Memo | 0 | yes |  |
| MatterID | Long | 4 | no |  |
| CaseID | Long | 4 | no |  |
| Date2 | DateTime | 8 | yes |  |
| Pay_Outlay | Memo | 0 | yes |  |
| Charge | Currency | 8 | yes |  |
| Payment | Currency | 8 | yes |  |
| FirmPrepaid | Boolean | 1 | yes |  |
| InsertPymt | Boolean | 1 | yes |  |
| AdvancedLegal | Boolean | 1 | yes |  |
| SSMA_TimeStamp | Binary | 8 | no |  |
| Orig_Atty | Text | 255 | yes |  |
| Case_Letter | Text | 255 | yes |  |
| CodeVal | Text | 255 | yes |  |
| Creimb | Boolean | 1 | yes |  |
| Closed | Boolean | 1 | yes |  |

**Primary Key:** CaseID

**Linked Connect:** `ODBC;DRIVER=SQL Server;SERVER=awsql2022dev;APP=Microsoft Office;DATABASE=TateBywater`

**Indexes:**
- UniqueIndex [primary, unique]: CaseID

### vw_advanced_totals_SUM (linked)
| Column | Type | Size | Nullable | Default |
|--------|------|------|----------|---------|
| CaseID | Long | 4 | no |  |
| SumOfCharge | Currency | 8 | yes |  |
| FirmPrepaid | Boolean | 1 | yes |  |

**Primary Key:** CaseID

**Linked Connect:** `ODBC;DRIVER=SQL Server;SERVER=awsql2022dev;APP=Microsoft Office;DATABASE=TateBywater`

**Indexes:**
- UniqueIndex [primary, unique]: CaseID

### vw_current_invoice (linked)
| Column | Type | Size | Nullable | Default |
|--------|------|------|----------|---------|
| Balance | Currency | 8 | yes |  |
| Balance Due Date | DateTime | 8 | yes |  |
| Billing Notes | Memo | 0 | yes |  |
| Last_Name | Text | 255 | yes |  |
| First_Name | Text | 255 | yes |  |
| CaseOpenDate | DateTime | 8 | yes |  |
| yr | Text | 254 | yes |  |
| Number_ | Long | 4 | yes |  |
| Orig_Atty | Text | 255 | yes |  |
| Address | Text | 255 | yes |  |
| City | Text | 255 | yes |  |
| State | Text | 255 | yes |  |
| Zip | Text | 255 | yes |  |
| Matter_type | Text | 255 | yes |  |
| Retainer | Currency | 8 | yes |  |
| Case_Letter | Text | 255 | yes |  |
| OrderNr | Long | 4 | yes |  |
| CaseID | Long | 4 | no |  |
| Date2 | DateTime | 8 | yes |  |
| Pay_Outlay | Memo | 0 | yes |  |
| Charge | Currency | 8 | yes |  |
| Payment | Currency | 8 | yes |  |

**Primary Key:** CaseID

**Linked Connect:** `ODBC;DRIVER=SQL Server;SERVER=awsql2022dev;APP=Microsoft Office;DATABASE=TateBywater`

**Indexes:**
- UniqueIndex [primary, unique]: CaseID

### vw_frm_invoices_summary (linked)
| Column | Type | Size | Nullable | Default |
|--------|------|------|----------|---------|
| CaseID | Long | 4 | yes |  |
| Name | Memo | 0 | no |  |
| First_Name | Text | 255 | yes |  |
| Last_Name | Text | 255 | yes |  |
| Closed | Boolean | 1 | yes |  |
| Retainer | Currency | 8 | yes |  |
| SumOfCharge | Currency | 8 | yes |  |
| SumOfPayment | Currency | 8 | yes |  |
| SumOfBalance | Currency | 8 | yes |  |
| BalanceCalculated | Currency | 8 | yes |  |
| BalRetCalculated | Currency | 8 | yes |  |
| FileNumber | Memo | 0 | yes |  |
| Balance Due Date | DateTime | 8 | yes |  |
| Orig_Atty | Text | 255 | yes |  |
| HandlingAtty_Case | Text | 255 | yes |  |
| CodeVal | Text | 255 | yes |  |
| Executor | Text | 255 | yes |  |
| LastOfInvSent | DateTime | 8 | yes |  |
| Long Term Collections | Boolean | 1 | yes |  |

**Primary Key:** CaseID

**Linked Connect:** `ODBC;DRIVER=SQL Server;SERVER=awsql2022dev;APP=Microsoft Office;DATABASE=TateBywater`

**Indexes:**
- UniqueIndex [primary, unique]: CaseID

### vw_invoice_comprehensive_trust_acc_cur_unfiltered (linked)
| Column | Type | Size | Nullable | Default |
|--------|------|------|----------|---------|
| CaseID | Long | 4 | yes |  |
| OrderNr | Long | 4 | yes |  |
| TDate | DateTime | 8 | yes |  |
| TMatter | Memo | 0 | yes |  |
| Debit | Currency | 8 | yes |  |
| balance | Currency | 8 | yes |  |

**Primary Key:** CaseID

**Linked Connect:** `ODBC;DRIVER=SQL Server;SERVER=awsql2022dev;APP=Microsoft Office;DATABASE=TateBywater`

**Indexes:**
- UniqueIndex [primary, unique]: CaseID

### vw_invoices_summary (linked)
| Column | Type | Size | Nullable | Default |
|--------|------|------|----------|---------|
| CaseID | Long | 4 | yes |  |
| Name | Memo | 0 | no |  |
| First_Name | Text | 255 | yes |  |
| Last_Name | Text | 255 | yes |  |
| Closed | Boolean | 1 | yes |  |
| Retainer | Currency | 8 | yes |  |
| SumOfCharge | Currency | 8 | yes |  |
| SumOfPayment | Currency | 8 | yes |  |
| SumOfBalance | Currency | 8 | yes |  |
| BalanceCalculated | Currency | 8 | yes |  |
| BalRetCalculated | Currency | 8 | yes |  |
| FileNumber | Memo | 0 | yes |  |
| Balance Due Date | DateTime | 8 | yes |  |
| Orig_Atty | Text | 255 | yes |  |
| HandlingAtty_Case | Text | 255 | yes |  |
| CodeVal | Text | 255 | yes |  |
| Executor | Text | 255 | yes |  |
| Long Term Collections | Boolean | 1 | yes |  |

**Primary Key:** CaseID

**Linked Connect:** `ODBC;DRIVER=SQL Server;SERVER=awsql2022dev;APP=Microsoft Office;DATABASE=TateBywater`

**Indexes:**
- UniqueIndex [primary, unique]: CaseID

### vw_last_invoice_sent (linked)
| Column | Type | Size | Nullable | Default |
|--------|------|------|----------|---------|
| CaseID | Long | 4 | yes |  |
| LastOfInvSent | DateTime | 8 | yes |  |

**Primary Key:** CaseID

**Linked Connect:** `ODBC;DRIVER=SQL Server;SERVER=awsql2022dev;APP=Microsoft Office;DATABASE=TateBywater`

**Indexes:**
- UniqueIndex [primary, unique]: CaseID

### vw_max_matterID_by_orderNr (linked)
| Column | Type | Size | Nullable | Default |
|--------|------|------|----------|---------|
| MaxOfOrderNr | Long | 4 | yes |  |
| MaxOfMatterID | Long | 4 | yes |  |
| CaseID | Long | 4 | no |  |

**Primary Key:** CaseID

**Linked Connect:** `ODBC;DRIVER=SQL Server;SERVER=awsql2022dev;APP=Microsoft Office;DATABASE=TateBywater`

**Indexes:**
- UniqueIndex [primary, unique]: CaseID

### vw_rpt_Matter_Closing (linked)
| Column | Type | Size | Nullable | Default |
|--------|------|------|----------|---------|
| CaseID | Long | 4 | no |  |
| MatterID | Long | 4 | no |  |
| Date2 | DateTime | 8 | yes |  |
| Pay_Outlay | Memo | 0 | yes |  |
| Charge | Currency | 8 | yes |  |
| Payment | Currency | 8 | yes |  |
| Balance | Currency | 8 | yes |  |
| RunningDebit | Currency | 8 | yes |  |
| RunningCredit | Currency | 8 | yes |  |
| RunningBalance | Currency | 8 | yes |  |
| Retainer | Currency | 8 | yes |  |
| RetBal | Currency | 8 | yes |  |
| OrderNr | Long | 4 | yes |  |

**Primary Key:** MatterID

**Linked Connect:** `ODBC;DRIVER=SQL Server;SERVER=awsql2022dev;APP=Microsoft Office;DATABASE=TateBywater`

**Indexes:**
- UniqueIndex [primary, unique]: MatterID

### vw_take_off_step2_sums (linked)
| Column | Type | Size | Nullable | Default |
|--------|------|------|----------|---------|
| TakeOffMonthID | Long | 4 | no |  |
| SumOfCBHRev | Currency | 8 | yes |  |
| SumOfMKRev | Currency | 8 | yes |  |
| SumOfCBHCom | Currency | 8 | yes |  |
| SumOfMTRev | Currency | 8 | yes |  |
| SumOfMTCom | Currency | 8 | yes |  |
| SumOfKBCom | Currency | 8 | yes |  |
| SumOfMKCom | Currency | 8 | yes |  |
| SumOfRLFCom | Currency | 8 | yes |  |
| SumOfEarlyEarned | Currency | 8 | yes |  |
| SumOfTOEarned | Currency | 8 | yes |  |
| SumOfTOEarlyAndEarned | Currency | 8 | yes |  |
| SumOfCostReimb | Currency | 8 | yes |  |

**Primary Key:** TakeOffMonthID

**Linked Connect:** `ODBC;DRIVER=SQL Server;SERVER=awsql2022dev;APP=Microsoft Office;DATABASE=TateBywater`

**Indexes:**
- UniqueIndex [primary, unique]: TakeOffMonthID

### vw_time_table_totals_atty_SUM (linked)
| Column | Type | Size | Nullable | Default |
|--------|------|------|----------|---------|
| SumOfSumOfAmount | Double | 8 | yes |  |
| Tatty | Text | 255 | yes |  |
| CaseID | Long | 4 | yes |  |

**Primary Key:** CaseID

**Linked Connect:** `ODBC;DRIVER=SQL Server;SERVER=awsql2022dev;APP=Microsoft Office;DATABASE=TateBywater`

**Indexes:**
- UniqueIndex [primary, unique]: CaseID

### vw_time_table_totals_SUM (linked)
| Column | Type | Size | Nullable | Default |
|--------|------|------|----------|---------|
| SumOfSumOfAmount | Double | 8 | yes |  |
| CaseID | Long | 4 | yes |  |

**Primary Key:** CaseID

**Linked Connect:** `ODBC;DRIVER=SQL Server;SERVER=awsql2022dev;APP=Microsoft Office;DATABASE=TateBywater`

**Indexes:**
- UniqueIndex [primary, unique]: CaseID

### vwAdvLegalFeesSum (linked)
| Column | Type | Size | Nullable | Default |
|--------|------|------|----------|---------|
| CaseID | Long | 4 | no |  |
| SumOfCharge | Currency | 8 | yes |  |
| AdvancedLegal | Boolean | 1 | yes |  |

**Primary Key:** CaseID

**Linked Connect:** `ODBC;DRIVER=SQL Server;SERVER=awsql2022dev;APP=Microsoft Office;DATABASE=TateBywater`

**Indexes:**
- UniqueIndex [primary, unique]: CaseID

### vwARCreditsSum (linked)
| Column | Type | Size | Nullable | Default |
|--------|------|------|----------|---------|
| CaseID | Long | 4 | no |  |
| SumOfPayment | Currency | 8 | yes |  |

**Primary Key:** CaseID

**Linked Connect:** `ODBC;DRIVER=SQL Server;SERVER=awsql2022dev;APP=Microsoft Office;DATABASE=TateBywater`

**Indexes:**
- UniqueIndex [primary, unique]: CaseID

### vwBillingTracker2 (linked)
| Column | Type | Size | Nullable | Default |
|--------|------|------|----------|---------|
| Bill_ID | Long | 4 | yes |  |
| Tdate | DateTime | 8 | yes |  |
| Tatty | Text | 255 | yes |  |
| Rate | Currency | 8 | yes |  |
| Time_ | Double | 8 | yes |  |
| Billed | Double | 8 | yes |  |
| CaseID | Long | 4 | yes |  |
| Last_Name | Text | 255 | yes |  |
| First_Name | Text | 255 | yes |  |
| Case_Letter | Text | 255 | yes |  |
| yr | Text | 254 | yes |  |
| Number_ | Long | 4 | yes |  |
| Orig_Atty | Text | 255 | yes |  |
| Name | Memo | 0 | yes |  |
| FileNumber | Memo | 0 | yes |  |
| Time_ID | Long | 4 | no |  |

**Primary Key:** Time_ID

**Linked Connect:** `ODBC;DRIVER=SQL Server;SERVER=awsql2022dev;APP=Microsoft Office;DATABASE=TateBywater`

**Indexes:**
- UniqueIndex [primary, unique]: Time_ID

### vwCaseListOpen (linked)
| Column | Type | Size | Nullable | Default |
|--------|------|------|----------|---------|
| CaseID | Long | 4 | no |  |
| CaseOpenDate | DateTime | 8 | yes |  |
| ClientName | Memo | 0 | no |  |
| Case_Letter | Text | 255 | yes |  |
| yr | Text | 254 | yes |  |
| Number_ | Long | 4 | yes |  |
| Orig_Atty | Text | 255 | yes |  |
| Extended_Ledger | Text | 255 | yes |  |
| Court | Text | 255 | yes |  |
| Matter_type | Text | 255 | yes |  |
| FileNumber | Memo | 0 | yes |  |
| Scan Location | Memo | 0 | yes |  |
| HandlingAtty_Case | Text | 255 | yes |  |
| Closed | Boolean | 1 | yes |  |
| CodeVal | Text | 255 | yes |  |
| ParaLegal | Text | 255 | yes |  |
| Retainer | Currency | 8 | yes |  |
| PIStatus | Text | 50 | no |  |

**Primary Key:** CaseID

**Linked Connect:** `ODBC;DRIVER=SQL Server;SERVER=awsql2022dev;APP=Microsoft Office;DATABASE=TateBywater`

**Indexes:**
- UniqueIndex [primary, unique]: CaseID

### vwCaseSourcesRPT (linked)
| Column | Type | Size | Nullable | Default |
|--------|------|------|----------|---------|
| CaseID | Long | 4 | no |  |
| Case_Letter | Text | 255 | yes |  |
| Number_ | Long | 4 | yes |  |
| Orig_Atty | Text | 255 | yes |  |
| Matter_type | Text | 255 | yes |  |
| CaseOpenDate | DateTime | 8 | yes |  |
| CaseNo | Memo | 0 | yes |  |
| yr | Text | 254 | yes |  |
| Total Earned Fee | Currency | 8 | yes |  |
| Clsdate | DateTime | 8 | yes |  |
| Closed | Boolean | 1 | yes |  |

**Primary Key:** CaseID

**Linked Connect:** `ODBC;DRIVER=SQL Server;SERVER=awsql2022dev;APP=Microsoft Office;DATABASE=TateBywater`

**Indexes:**
- UniqueIndex [primary, unique]: CaseID

### vwCostReimbSUM (linked)
| Column | Type | Size | Nullable | Default |
|--------|------|------|----------|---------|
| CaseID | Long | 4 | yes |  |
| SumOfCredit | Currency | 8 | yes |  |

**Primary Key:** CaseID

**Linked Connect:** `ODBC;DRIVER=SQL Server;SERVER=awsql2022dev;APP=Microsoft Office;DATABASE=TateBywater`

**Indexes:**
- UniqueIndex [primary, unique]: CaseID

### vwDispos (linked)
| Column | Type | Size | Nullable | Default |
|--------|------|------|----------|---------|
| CaseID | Long | 4 | no |  |
| Case_Letter | Text | 255 | yes |  |
| Orig_Atty | Text | 255 | yes |  |
| Matter_type | Text | 255 | yes |  |
| Court | Text | 255 | yes |  |
| CaseOpenDate | DateTime | 8 | yes |  |
| HandlingAtty_Case | Text | 255 | yes |  |
| Dispo_Atty | Text | 255 | yes |  |
| Dispo_Date | DateTime | 8 | yes |  |
| PI Settlement Amount | Currency | 8 | yes |  |
| Entire np | Boolean | 1 | yes |  |
| Not Guilty Dismissed | Boolean | 1 | yes |  |
| Plea | Boolean | 1 | yes |  |
| Trial | Boolean | 1 | yes |  |
| Disposition | Memo | 0 | yes |  |
| Name | Memo | 0 | yes |  |
| Case No | Memo | 0 | yes |  |
| Litigation | Boolean | 1 | yes |  |
| CodeVal | Text | 255 | yes |  |
| FieldName | Text | 255 | yes |  |

**Primary Key:** CaseID

**Linked Connect:** `ODBC;DRIVER=SQL Server;SERVER=awsql2022dev;APP=Microsoft Office;DATABASE=TateBywater`

**Indexes:**
- UniqueIndex [primary, unique]: CaseID

### vwEarnedAdvLegalSUM (linked)
| Column | Type | Size | Nullable | Default |
|--------|------|------|----------|---------|
| CaseID | Long | 4 | yes |  |
| SumOfCredit | Currency | 8 | yes |  |

**Primary Key:** CaseID

**Linked Connect:** `ODBC;DRIVER=SQL Server;SERVER=awsql2022dev;APP=Microsoft Office;DATABASE=TateBywater`

**Indexes:**
- UniqueIndex [primary, unique]: CaseID

### vwfrmClientLedger (linked)
| Column | Type | Size | Nullable | Default |
|--------|------|------|----------|---------|
| CaseID | Long | 4 | no |  |
| Last_Name | Text | 255 | yes |  |
| First_Name | Text | 255 | yes |  |
| CaseOpenDate | DateTime | 8 | yes |  |
| Closed | Boolean | 1 | yes |  |
| Clsdate | DateTime | 8 | yes |  |
| Extended_Ledger | Text | 255 | yes |  |
| Case_Letter | Text | 255 | yes |  |
| yr | Text | 254 | yes |  |
| Number_ | Long | 4 | yes |  |
| Orig_Atty | Text | 255 | yes |  |
| Address | Text | 255 | yes |  |
| CourtCaseNo | Text | 255 | yes |  |
| City | Text | 255 | yes |  |
| FamilyLaw | Boolean | 1 | yes |  |
| State | Text | 255 | yes |  |
| Zip | Text | 255 | yes |  |
| Country | Text | 255 | yes |  |
| HmPhone | Text | 255 | yes |  |
| Action | Text | 255 | yes |  |
| OtherPhone | Text | 255 | yes |  |
| Fax | Text | 255 | yes |  |
| WkPhone | Text | 255 | yes |  |
| Comments | Memo | 0 | yes |  |
| Email | Text | 255 | yes |  |
| Referral | Text | 255 | yes |  |
| Individual Referrer | Text | 255 | yes |  |
| Retainer | Currency | 8 | yes |  |
| Matter_type | Text | 255 | yes |  |
| SOL | DateTime | 8 | yes |  |
| Court | Text | 255 | yes |  |
| CType | Text | 255 | yes |  |
| POfc | Text | 255 | yes |  |
| ComplainingWitness | Text | 255 | yes |  |
| DOB | DateTime | 8 | yes |  |
| WkAddress | Text | 255 | yes |  |
| WkCity | Text | 255 | yes |  |
| WkState | Text | 255 | yes |  |
| WkZip | Text | 255 | yes |  |
| Pro Bono | Boolean | 1 | yes |  |
| HandlingAtty_Case | Text | 255 | yes |  |
| Action_Needed_on_Payment | Boolean | 1 | yes |  |
| SSN | Text | 255 | yes |  |
| Employer Name | Text | 255 | yes |  |
| Last Updated Contact Info | DateTime | 8 | yes |  |
| Ocounsel | Text | 255 | yes |  |
| Firm | Text | 255 | yes |  |
| OC Address | Text | 255 | yes |  |
| OC City | Text | 255 | yes |  |
| OC State | Text | 255 | yes |  |
| OC Zip | Text | 255 | yes |  |
| OC Phone | Text | 255 | yes |  |
| OC Email | Text | 255 | yes |  |
| OC Fax | Text | 255 | yes |  |
| Pro Bono PM | Text | 255 | yes |  |
| Pro Bono JRT | Text | 255 | yes |  |
| ContingencyFee | Boolean | 1 | yes |  |
| AuthorityToTalkTo | Memo | 0 | yes |  |
| Hourly | Boolean | 1 | yes |  |
| Contingency | Boolean | 1 | yes |  |
| Hybrid | Boolean | 1 | yes |  |
| Family-Law | Boolean | 1 | yes |  |
| Fixed | Boolean | 1 | yes |  |
| Scan | Boolean | 1 | yes |  |
| Scan Location | Memo | 0 | yes |  |
| ScanNotAvail | Boolean | 1 | yes |  |
| ParaLegal | Text | 255 | yes |  |
| Spanish | Boolean | 1 | yes |  |
| Offdate | DateTime | 8 | yes |  |
| CostHold | Currency | 8 | yes |  |
| CltNarrative | Memo | 0 | yes |  |
| ARTrustZero | Boolean | 1 | yes |  |
| F73 | Text | 255 | yes |  |
| F74 | Text | 255 | yes |  |
| F75 | Text | 255 | yes |  |
| F76 | Text | 255 | yes |  |
| PhName1 | Text | 255 | yes |  |
| PhName2 | Text | 255 | yes |  |
| LengthRes | Text | 255 | yes |  |
| LengthEmp | Text | 255 | yes |  |
| LegalStatus | Text | 255 | yes |  |
| CurrentBond | Text | 255 | yes |  |
| CrRecord | Memo | 0 | yes |  |
| TrustChronMemo | Memo | 0 | yes |  |
| Executor | Text | 255 | yes |  |
| RetainerReimb | Boolean | 1 | yes |  |
| RetReimbAmount | Currency | 8 | yes |  |
| Reviewable | Boolean | 1 | yes |  |
| ReviewReq | DateTime | 8 | yes |  |
| ReviewReceivedDate | DateTime | 8 | yes |  |
| ReviewReceived | Boolean | 1 | yes |  |
| Testimonial | Memo | 0 | yes |  |
| ReviewFollowUp | DateTime | 8 | yes |  |
| Stars | Long | 4 | yes |  |
| Review Source | Text | 255 | yes |  |
| Review Date | DateTime | 8 | yes |  |
| Title | Text | 255 | yes |  |
| OPartyLast | Text | 255 | yes |  |
| OPartyFirst | Text | 255 | yes |  |
| OPartyDOB | DateTime | 8 | yes |  |
| SSMA_TimeStamp | Binary | 8 | no |  |
| FileNo | Memo | 0 | yes |  |
| PartnerRate | Currency | 8 | yes |  |
| AssocRate | Currency | 8 | yes |  |

**Primary Key:** CaseID

**Linked Connect:** `ODBC;DRIVER=SQL Server;SERVER=awsql2022dev;APP=Microsoft Office;DATABASE=TateBywater`

**Indexes:**
- UniqueIndex [primary, unique]: CaseID

### vwfrmTakeOffSubForm (linked)
| Column | Type | Size | Nullable | Default |
|--------|------|------|----------|---------|
| FileNumber | Memo | 0 | yes |  |
| Name | Memo | 0 | no |  |
| CaseID | Long | 4 | no |  |
| Last_Name | Text | 255 | yes |  |
| First_Name | Text | 255 | yes |  |
| CaseOpenDate | DateTime | 8 | yes |  |
| Closed | Boolean | 1 | yes |  |
| Clsdate | DateTime | 8 | yes |  |
| Extended_Ledger | Text | 255 | yes |  |
| Case_Letter | Text | 255 | yes |  |
| yr | Text | 254 | yes |  |
| Number_ | Long | 4 | yes |  |
| Orig_Atty | Text | 255 | yes |  |
| Address | Text | 255 | yes |  |
| CourtCaseNo | Text | 255 | yes |  |
| City | Text | 255 | yes |  |
| FamilyLaw | Boolean | 1 | yes |  |
| State | Text | 255 | yes |  |
| Zip | Text | 255 | yes |  |
| Country | Text | 255 | yes |  |
| HmPhone | Text | 255 | yes |  |
| Action | Text | 255 | yes |  |
| OtherPhone | Text | 255 | yes |  |
| Fax | Text | 255 | yes |  |
| WkPhone | Text | 255 | yes |  |
| Comments | Memo | 0 | yes |  |
| Email | Text | 255 | yes |  |
| Referral | Text | 255 | yes |  |
| Individual Referrer | Text | 255 | yes |  |
| Retainer | Currency | 8 | yes |  |
| Matter_type | Text | 255 | yes |  |
| SOL | DateTime | 8 | yes |  |
| Court | Text | 255 | yes |  |
| CType | Text | 255 | yes |  |
| POfc | Text | 255 | yes |  |
| ComplainingWitness | Text | 255 | yes |  |
| DOB | DateTime | 8 | yes |  |
| WkAddress | Text | 255 | yes |  |
| WkCity | Text | 255 | yes |  |
| WkState | Text | 255 | yes |  |
| WkZip | Text | 255 | yes |  |
| Pro Bono | Boolean | 1 | yes |  |
| HandlingAtty_Case | Text | 255 | yes |  |
| Action_Needed_on_Payment | Boolean | 1 | yes |  |
| SSN | Text | 255 | yes |  |
| Employer Name | Text | 255 | yes |  |
| Last Updated Contact Info | DateTime | 8 | yes |  |
| Ocounsel | Text | 255 | yes |  |
| Firm | Text | 255 | yes |  |
| OC Address | Text | 255 | yes |  |
| OC City | Text | 255 | yes |  |
| OC State | Text | 255 | yes |  |
| OC Zip | Text | 255 | yes |  |
| OC Phone | Text | 255 | yes |  |
| OC Email | Text | 255 | yes |  |
| OC Fax | Text | 255 | yes |  |
| Pro Bono PM | Text | 255 | yes |  |
| Pro Bono JRT | Text | 255 | yes |  |
| ContingencyFee | Boolean | 1 | yes |  |
| AuthorityToTalkTo | Memo | 0 | yes |  |
| Hourly | Boolean | 1 | yes |  |
| Contingency | Boolean | 1 | yes |  |
| Hybrid | Boolean | 1 | yes |  |
| Family-Law | Boolean | 1 | yes |  |
| Fixed | Boolean | 1 | yes |  |
| Scan | Boolean | 1 | yes |  |
| Scan Location | Memo | 0 | yes |  |
| ScanNotAvail | Boolean | 1 | yes |  |
| ParaLegal | Text | 255 | yes |  |
| Spanish | Boolean | 1 | yes |  |
| Offdate | DateTime | 8 | yes |  |
| CostHold | Currency | 8 | yes |  |
| CltNarrative | Memo | 0 | yes |  |
| ARTrustZero | Boolean | 1 | yes |  |
| F73 | Text | 255 | yes |  |
| F74 | Text | 255 | yes |  |
| F75 | Text | 255 | yes |  |
| F76 | Text | 255 | yes |  |
| PhName1 | Text | 255 | yes |  |
| PhName2 | Text | 255 | yes |  |
| LengthRes | Text | 255 | yes |  |
| LengthEmp | Text | 255 | yes |  |
| LegalStatus | Text | 255 | yes |  |
| CurrentBond | Text | 255 | yes |  |
| CrRecord | Memo | 0 | yes |  |
| TrustChronMemo | Memo | 0 | yes |  |
| Executor | Text | 255 | yes |  |
| RetainerReimb | Boolean | 1 | yes |  |
| RetReimbAmount | Currency | 8 | yes |  |
| Reviewable | Boolean | 1 | yes |  |
| ReviewReq | DateTime | 8 | yes |  |
| ReviewReceivedDate | DateTime | 8 | yes |  |
| ReviewReceived | Boolean | 1 | yes |  |
| Testimonial | Memo | 0 | yes |  |
| ReviewFollowUp | DateTime | 8 | yes |  |
| Stars | Long | 4 | yes |  |
| Review Source | Text | 255 | yes |  |
| Review Date | DateTime | 8 | yes |  |
| Title | Text | 255 | yes |  |
| OPartyLast | Text | 255 | yes |  |
| OPartyFirst | Text | 255 | yes |  |
| OPartyDOB | DateTime | 8 | yes |  |
| TakeOffID | Long | 4 | yes |  |
| TakeOffMonthID | Long | 4 | yes |  |
| AvailBalance | Currency | 8 | yes |  |
| TotalUnCashedChks | Currency | 8 | yes |  |
| TotalUnclearedDeps | Currency | 8 | yes |  |
| TotalAdvancedAR | Currency | 8 | yes |  |
| EarlyEarned | Currency | 8 | yes |  |
| TOEarned | Currency | 8 | yes |  |
| TOAttBilled | Currency | 8 | yes |  |
| CostReimb | Currency | 8 | yes |  |
| CBHRev | Currency | 8 | yes |  |
| MKRev | Currency | 8 | yes |  |
| CBHCom | Currency | 8 | yes |  |
| MTRev | Currency | 8 | yes |  |
| MTCom | Currency | 8 | yes |  |
| KBCom | Currency | 8 | yes |  |
| MKCom | Currency | 8 | yes |  |
| TOEarnedTr | Boolean | 1 | yes |  |
| CostReimbTr | Boolean | 1 | yes |  |
| InsertedTrust | Boolean | 1 | yes |  |
| TotalHourlyOuts | Currency | 8 | yes |  |
| OpenTK | Text | 20 | yes |  |
| AdvCostBal | Currency | 8 | yes |  |
| AdvFeeBal | Currency | 8 | yes |  |
| CostHoldBal | Currency | 8 | yes |  |
| BRRev | Currency | 8 | yes |  |
| BRCom | Currency | 8 | yes |  |
| RLFCom | Currency | 8 | yes |  |
| AdvEarned | Currency | 8 | no |  |
| RemEarned | Currency | 8 | yes |  |
| SumOfCBHRev | Currency | 8 | yes |  |
| SumOfMKRev | Currency | 8 | yes |  |
| SumOfCBHCom | Currency | 8 | yes |  |
| SumOfMTRev | Currency | 8 | yes |  |
| SumOfMTCom | Currency | 8 | yes |  |
| SumOfKBCom | Currency | 8 | yes |  |
| SumOfMKCom | Currency | 8 | yes |  |
| SumOfRLFCom | Currency | 8 | yes |  |
| SumOfEarlyEarned | Currency | 8 | yes |  |
| SumOfTOEarned | Currency | 8 | yes |  |
| SumOfTOEarlyAndEarned | Currency | 8 | yes |  |
| SumOfCostReimb | Currency | 8 | yes |  |

**Primary Key:** TakeOffID

**Linked Connect:** `ODBC;DRIVER=SQL Server;SERVER=awsql2022dev;APP=Microsoft Office;DATABASE=TateBywater`

**Indexes:**
- UniqueIndex [primary, unique]: TakeOffID

### vwInvoiceComprehensiveTrust (linked)
| Column | Type | Size | Nullable | Default |
|--------|------|------|----------|---------|
| TrustAccountID | Long | 4 | yes |  |
| TDate | DateTime | 8 | yes |  |
| TMatter | Memo | 0 | yes |  |
| Debit | Currency | 8 | yes |  |
| Credit | Currency | 8 | yes |  |
| CaseID | Long | 4 | yes |  |

**Primary Key:** CaseID

**Linked Connect:** `ODBC;DRIVER=SQL Server;SERVER=awsql2022dev;APP=Microsoft Office;DATABASE=TateBywater`

**Indexes:**
- UniqueIndex [primary, unique]: CaseID

### vwInvoiceRPT (linked)
| Column | Type | Size | Nullable | Default |
|--------|------|------|----------|---------|
| CaseID | Long | 4 | yes |  |
| MatterID | Long | 4 | yes |  |
| Date2 | DateTime | 8 | yes |  |
| Pay_Outlay | Memo | 0 | yes |  |
| Charge | Currency | 8 | yes |  |
| Payment | Currency | 8 | yes |  |
| Case No | Memo | 0 | yes |  |
| ID | Long | 4 | yes |  |
| Balance Due Date | DateTime | 8 | yes |  |
| Past Due | Boolean | 1 | yes |  |
| Long Term Collections | Boolean | 1 | yes |  |
| chkBalanceDue | Boolean | 1 | yes |  |
| Billing Notes | Memo | 0 | yes |  |

**Primary Key:** CaseID

**Linked Connect:** `ODBC;DRIVER=SQL Server;SERVER=awsql2022dev;APP=Microsoft Office;DATABASE=TateBywater`

**Indexes:**
- UniqueIndex [primary, unique]: CaseID

### vwInvoiceRPT1 (linked)
| Column | Type | Size | Nullable | Default |
|--------|------|------|----------|---------|
| CaseID | Long | 4 | yes |  |
| Last_Name | Text | 255 | yes |  |
| First_Name | Text | 255 | yes |  |
| CaseOpenDate | DateTime | 8 | yes |  |
| Closed | Boolean | 1 | yes |  |
| Clsdate | DateTime | 8 | yes |  |
| Extended_Ledger | Text | 255 | yes |  |
| Case_Letter | Text | 255 | yes |  |
| yr | Text | 254 | yes |  |
| Number_ | Long | 4 | yes |  |
| Orig_Atty | Text | 255 | yes |  |
| Address | Text | 255 | yes |  |
| CourtCaseNo | Text | 255 | yes |  |
| City | Text | 255 | yes |  |
| FamilyLaw | Boolean | 1 | yes |  |
| State | Text | 255 | yes |  |
| Zip | Text | 255 | yes |  |
| Country | Text | 255 | yes |  |
| HmPhone | Text | 255 | yes |  |
| Action | Text | 255 | yes |  |
| OtherPhone | Text | 255 | yes |  |
| Fax | Text | 255 | yes |  |
| WkPhone | Text | 255 | yes |  |
| Comments | Memo | 0 | yes |  |
| Email | Text | 255 | yes |  |
| Referral | Text | 255 | yes |  |
| Individual Referrer | Text | 255 | yes |  |
| Retainer | Currency | 8 | yes |  |
| Matter_type | Text | 255 | yes |  |
| SOL | DateTime | 8 | yes |  |
| Court | Text | 255 | yes |  |
| CType | Text | 255 | yes |  |
| POfc | Text | 255 | yes |  |
| ComplainingWitness | Text | 255 | yes |  |
| DOB | DateTime | 8 | yes |  |
| WkAddress | Text | 255 | yes |  |
| WkCity | Text | 255 | yes |  |
| WkState | Text | 255 | yes |  |
| WkZip | Text | 255 | yes |  |
| Pro Bono | Boolean | 1 | yes |  |
| HandlingAtty_Case | Text | 255 | yes |  |
| Action_Needed_on_Payment | Boolean | 1 | yes |  |
| SSN | Text | 255 | yes |  |
| Employer Name | Text | 255 | yes |  |
| Last Updated Contact Info | DateTime | 8 | yes |  |
| Ocounsel | Text | 255 | yes |  |
| Firm | Text | 255 | yes |  |
| OC Address | Text | 255 | yes |  |
| OC City | Text | 255 | yes |  |
| OC State | Text | 255 | yes |  |
| OC Zip | Text | 255 | yes |  |
| OC Phone | Text | 255 | yes |  |
| OC Email | Text | 255 | yes |  |
| OC Fax | Text | 255 | yes |  |
| Pro Bono PM | Text | 255 | yes |  |
| Pro Bono JRT | Text | 255 | yes |  |
| ContingencyFee | Boolean | 1 | yes |  |
| AuthorityToTalkTo | Memo | 0 | yes |  |
| Hourly | Boolean | 1 | yes |  |
| Contingency | Boolean | 1 | yes |  |
| Hybrid | Boolean | 1 | yes |  |
| Family-Law | Boolean | 1 | yes |  |
| Fixed | Boolean | 1 | yes |  |
| Scan | Boolean | 1 | yes |  |
| Scan Location | Memo | 0 | yes |  |
| ScanNotAvail | Boolean | 1 | yes |  |
| ParaLegal | Text | 255 | yes |  |
| Spanish | Boolean | 1 | yes |  |
| Offdate | DateTime | 8 | yes |  |
| CostHold | Currency | 8 | yes |  |
| CltNarrative | Memo | 0 | yes |  |
| ARTrustZero | Boolean | 1 | yes |  |
| F73 | Text | 255 | yes |  |
| F74 | Text | 255 | yes |  |
| F75 | Text | 255 | yes |  |
| F76 | Text | 255 | yes |  |
| PhName1 | Text | 255 | yes |  |
| PhName2 | Text | 255 | yes |  |
| LengthRes | Text | 255 | yes |  |
| LengthEmp | Text | 255 | yes |  |
| LegalStatus | Text | 255 | yes |  |
| CurrentBond | Text | 255 | yes |  |
| CrRecord | Memo | 0 | yes |  |
| TrustChronMemo | Memo | 0 | yes |  |
| Executor | Text | 255 | yes |  |
| RetainerReimb | Boolean | 1 | yes |  |
| RetReimbAmount | Currency | 8 | yes |  |
| Reviewable | Boolean | 1 | yes |  |
| ReviewReq | DateTime | 8 | yes |  |
| ReviewReceivedDate | DateTime | 8 | yes |  |
| ReviewReceived | Boolean | 1 | yes |  |
| Testimonial | Memo | 0 | yes |  |
| ReviewFollowUp | DateTime | 8 | yes |  |
| Stars | Long | 4 | yes |  |
| Review Source | Text | 255 | yes |  |
| Review Date | DateTime | 8 | yes |  |
| Title | Text | 255 | yes |  |
| OPartyLast | Text | 255 | yes |  |
| OPartyFirst | Text | 255 | yes |  |
| OPartyDOB | DateTime | 8 | yes |  |
| SSMA_TimeStamp | Binary | 8 | yes |  |
| Balance | Currency | 8 | yes |  |
| MatterID | Long | 4 | yes |  |
| Date2 | DateTime | 8 | yes |  |
| Pay_Outlay | Memo | 0 | yes |  |
| Charge | Currency | 8 | yes |  |
| Payment | Currency | 8 | yes |  |
| Case No | Memo | 0 | yes |  |
| Balance Due Date | DateTime | 8 | yes |  |
| Past Due | Boolean | 1 | yes |  |
| Long Term Collections | Boolean | 1 | yes |  |
| chkBalanceDue | Boolean | 1 | yes |  |
| Billing Notes | Memo | 0 | yes |  |
| RunningDebit | Currency | 8 | yes |  |
| RunningCredit | Currency | 8 | yes |  |
| RunningBalance | Currency | 8 | yes |  |
| OrderNr | Long | 4 | yes |  |

**Primary Key:** CaseID

**Linked Connect:** `ODBC;DRIVER=SQL Server;SERVER=awsql2022dev;APP=Microsoft Office;DATABASE=TateBywater`

**Indexes:**
- UniqueIndex [primary, unique]: CaseID

### vwMatter (linked)
| Column | Type | Size | Nullable | Default |
|--------|------|------|----------|---------|
| MatterID | Long | 4 | no |  |
| Date2 | DateTime | 8 | yes |  |
| CaseID | Long | 4 | no |  |
| SumOfCharge | Currency | 8 | yes |  |
| SumOfPayment | Currency | 8 | yes |  |
| Balance | Currency | 8 | yes |  |
| OrderNr | Long | 4 | yes |  |

**Primary Key:** MatterID

**Linked Connect:** `ODBC;DRIVER=SQL Server;SERVER=awsql2022dev;APP=Microsoft Office;DATABASE=TateBywater`

**Indexes:**
- UniqueIndex [primary, unique]: MatterID

### vwMatterAndAR (linked)
| Column | Type | Size | Nullable | Default |
|--------|------|------|----------|---------|
| MatterID | Long | 4 | no |  |
| CaseID | Long | 4 | no |  |
| Date2 | DateTime | 8 | yes |  |
| Pay_Outlay | Memo | 0 | yes |  |
| Charge | Currency | 8 | yes |  |
| Payment | Currency | 8 | yes |  |
| Creimb | Boolean | 1 | yes |  |
| FirmPrepaid | Boolean | 1 | yes |  |
| OrderNr | Long | 4 | yes |  |
| InsertPymt | Boolean | 1 | yes |  |
| AdvancedLegal | Boolean | 1 | yes |  |
| SumOfCharge | Currency | 8 | yes |  |
| SumOfPayment | Currency | 8 | yes |  |
| Retainer | Currency | 8 | yes |  |
| Balance | Currency | 8 | yes |  |

**Primary Key:** MatterID

**Linked Connect:** `ODBC;DRIVER=SQL Server;SERVER=awsql2022dev;APP=Microsoft Office;DATABASE=TateBywater`

**Indexes:**
- UniqueIndex [primary, unique]: MatterID

### vwMatterBalanceTotals (linked)
| Column | Type | Size | Nullable | Default |
|--------|------|------|----------|---------|
| CaseID | Long | 4 | no |  |
| SumOfBalance | Currency | 8 | yes |  |

**Primary Key:** CaseID

**Linked Connect:** `ODBC;DRIVER=SQL Server;SERVER=awsql2022dev;APP=Microsoft Office;DATABASE=TateBywater`

**Indexes:**
- UniqueIndex [primary, unique]: CaseID

### vwNewTrustComp (linked)
| Column | Type | Size | Nullable | Default |
|--------|------|------|----------|---------|
| CaseID | Long | 4 | yes |  |
| Last_Name | Text | 255 | yes |  |
| First_Name | Text | 255 | yes |  |
| CaseOpenDate | DateTime | 8 | yes |  |
| Closed | Boolean | 1 | yes |  |
| Clsdate | DateTime | 8 | yes |  |
| Extended_Ledger | Text | 255 | yes |  |
| Case_Letter | Text | 255 | yes |  |
| yr | Text | 254 | yes |  |
| Number_ | Long | 4 | yes |  |
| Orig_Atty | Text | 255 | yes |  |
| Address | Text | 255 | yes |  |
| CourtCaseNo | Text | 255 | yes |  |
| City | Text | 255 | yes |  |
| FamilyLaw | Boolean | 1 | yes |  |
| State | Text | 255 | yes |  |
| Zip | Text | 255 | yes |  |
| Country | Text | 255 | yes |  |
| HmPhone | Text | 255 | yes |  |
| Action | Text | 255 | yes |  |
| OtherPhone | Text | 255 | yes |  |
| Fax | Text | 255 | yes |  |
| WkPhone | Text | 255 | yes |  |
| Comments | Memo | 0 | yes |  |
| Email | Text | 255 | yes |  |
| Referral | Text | 255 | yes |  |
| Individual Referrer | Text | 255 | yes |  |
| Retainer | Currency | 8 | yes |  |
| Matter_type | Text | 255 | yes |  |
| SOL | DateTime | 8 | yes |  |
| Court | Text | 255 | yes |  |
| CType | Text | 255 | yes |  |
| POfc | Text | 255 | yes |  |
| ComplainingWitness | Text | 255 | yes |  |
| DOB | DateTime | 8 | yes |  |
| WkAddress | Text | 255 | yes |  |
| WkCity | Text | 255 | yes |  |
| WkState | Text | 255 | yes |  |
| WkZip | Text | 255 | yes |  |
| Pro Bono | Boolean | 1 | yes |  |
| HandlingAtty_Case | Text | 255 | yes |  |
| Action_Needed_on_Payment | Boolean | 1 | yes |  |
| SSN | Text | 255 | yes |  |
| Employer Name | Text | 255 | yes |  |
| Last Updated Contact Info | DateTime | 8 | yes |  |
| Ocounsel | Text | 255 | yes |  |
| Firm | Text | 255 | yes |  |
| OC Address | Text | 255 | yes |  |
| OC City | Text | 255 | yes |  |
| OC State | Text | 255 | yes |  |
| OC Zip | Text | 255 | yes |  |
| OC Phone | Text | 255 | yes |  |
| OC Email | Text | 255 | yes |  |
| OC Fax | Text | 255 | yes |  |
| Pro Bono PM | Text | 255 | yes |  |
| Pro Bono JRT | Text | 255 | yes |  |
| ContingencyFee | Boolean | 1 | yes |  |
| AuthorityToTalkTo | Memo | 0 | yes |  |
| Hourly | Boolean | 1 | yes |  |
| Contingency | Boolean | 1 | yes |  |
| Hybrid | Boolean | 1 | yes |  |
| Family-Law | Boolean | 1 | yes |  |
| Fixed | Boolean | 1 | yes |  |
| Scan | Boolean | 1 | yes |  |
| Scan Location | Memo | 0 | yes |  |
| ScanNotAvail | Boolean | 1 | yes |  |
| ParaLegal | Text | 255 | yes |  |
| Spanish | Boolean | 1 | yes |  |
| Offdate | DateTime | 8 | yes |  |
| CostHold | Currency | 8 | yes |  |
| CltNarrative | Memo | 0 | yes |  |
| ARTrustZero | Boolean | 1 | yes |  |
| F73 | Text | 255 | yes |  |
| F74 | Text | 255 | yes |  |
| F75 | Text | 255 | yes |  |
| F76 | Text | 255 | yes |  |
| PhName1 | Text | 255 | yes |  |
| PhName2 | Text | 255 | yes |  |
| LengthRes | Text | 255 | yes |  |
| LengthEmp | Text | 255 | yes |  |
| LegalStatus | Text | 255 | yes |  |
| CurrentBond | Text | 255 | yes |  |
| CrRecord | Memo | 0 | yes |  |
| TrustChronMemo | Memo | 0 | yes |  |
| Executor | Text | 255 | yes |  |
| RetainerReimb | Boolean | 1 | yes |  |
| RetReimbAmount | Currency | 8 | yes |  |
| Reviewable | Boolean | 1 | yes |  |
| ReviewReq | DateTime | 8 | yes |  |
| ReviewReceivedDate | DateTime | 8 | yes |  |
| ReviewReceived | Boolean | 1 | yes |  |
| Testimonial | Memo | 0 | yes |  |
| ReviewFollowUp | DateTime | 8 | yes |  |
| Stars | Long | 4 | yes |  |
| Review Source | Text | 255 | yes |  |
| Review Date | DateTime | 8 | yes |  |
| Title | Text | 255 | yes |  |
| OPartyLast | Text | 255 | yes |  |
| OPartyFirst | Text | 255 | yes |  |
| OPartyDOB | DateTime | 8 | yes |  |
| SSMA_TimeStamp | Binary | 8 | yes |  |
| SumOfDebit | Currency | 8 | yes |  |
| SumOfCredit | Currency | 8 | yes |  |
| Balance | Currency | 8 | yes |  |
| TrustAccountID | Long | 4 | yes |  |
| TDate | DateTime | 8 | yes |  |
| TMatter | Memo | 0 | yes |  |
| Debit | Currency | 8 | yes |  |
| Credit | Currency | 8 | yes |  |
| CheckNumber | Text | 255 | yes |  |
| Case No | Memo | 0 | yes |  |
| CheckCashed | Boolean | 1 | yes |  |
| DepCleared | Boolean | 1 | yes |  |
| Reconciled | Boolean | 1 | yes |  |
| OrderNr | Long | 4 | yes |  |
| Name | Memo | 0 | no |  |
| FileNumber | Memo | 0 | yes |  |

**Primary Key:** CaseID

**Linked Connect:** `ODBC;DRIVER=SQL Server;SERVER=awsql2022dev;APP=Microsoft Office;DATABASE=TateBywater`

**Indexes:**
- UniqueIndex [primary, unique]: CaseID

### vwPILogLatestDate (linked)
| Column | Type | Size | Nullable | Default |
|--------|------|------|----------|---------|
| MaxEventDate | DateTime | 8 | yes |  |
| ID | Long | 4 | yes |  |

**Primary Key:** ID

**Linked Connect:** `ODBC;DRIVER=SQL Server;SERVER=awsql2022dev;APP=Microsoft Office;DATABASE=TateBywater`

**Indexes:**
- UniqueIndex [primary, unique]: ID

### vwStmtTrustRPT (linked)
| Column | Type | Size | Nullable | Default |
|--------|------|------|----------|---------|
| CaseID | Long | 4 | no |  |
| Case_Letter | Text | 255 | yes |  |
| yr | Text | 254 | yes |  |
| Number_ | Long | 4 | yes |  |
| Orig_Atty | Text | 255 | yes |  |
| Matter_type | Text | 255 | yes |  |
| Retainer | Currency | 8 | yes |  |
| CaseOpenDate | DateTime | 8 | yes |  |
| TrustAccountID | Long | 4 | yes |  |
| TDate | DateTime | 8 | yes |  |
| TMatter | Memo | 0 | yes |  |
| Debit | Currency | 8 | yes |  |
| Credit | Currency | 8 | yes |  |
| CheckCashed | Boolean | 1 | yes |  |
| CheckNumber | Text | 255 | yes |  |
| Case No | Memo | 0 | yes |  |
| DepCleared | Boolean | 1 | yes |  |
| Reconciled | Boolean | 1 | yes |  |
| OrderNr | Long | 4 | yes |  |

**Primary Key:** CaseID

**Linked Connect:** `ODBC;DRIVER=SQL Server;SERVER=awsql2022dev;APP=Microsoft Office;DATABASE=TateBywater`

**Indexes:**
- UniqueIndex [primary, unique]: CaseID

### vwStmtTrustRPT1 (linked)
| Column | Type | Size | Nullable | Default |
|--------|------|------|----------|---------|
| CaseID | Long | 4 | yes |  |
| Last_Name | Text | 255 | yes |  |
| First_Name | Text | 255 | yes |  |
| CaseOpenDate | DateTime | 8 | yes |  |
| Closed | Boolean | 1 | yes |  |
| Clsdate | DateTime | 8 | yes |  |
| Extended_Ledger | Text | 255 | yes |  |
| Case_Letter | Text | 255 | yes |  |
| yr | Text | 254 | yes |  |
| Number_ | Long | 4 | yes |  |
| Orig_Atty | Text | 255 | yes |  |
| Address | Text | 255 | yes |  |
| CourtCaseNo | Text | 255 | yes |  |
| City | Text | 255 | yes |  |
| FamilyLaw | Boolean | 1 | yes |  |
| State | Text | 255 | yes |  |
| Zip | Text | 255 | yes |  |
| Country | Text | 255 | yes |  |
| HmPhone | Text | 255 | yes |  |
| Action | Text | 255 | yes |  |
| OtherPhone | Text | 255 | yes |  |
| Fax | Text | 255 | yes |  |
| WkPhone | Text | 255 | yes |  |
| Comments | Memo | 0 | yes |  |
| Email | Text | 255 | yes |  |
| Referral | Text | 255 | yes |  |
| Individual Referrer | Text | 255 | yes |  |
| Retainer | Currency | 8 | yes |  |
| Matter_type | Text | 255 | yes |  |
| SOL | DateTime | 8 | yes |  |
| Court | Text | 255 | yes |  |
| CType | Text | 255 | yes |  |
| POfc | Text | 255 | yes |  |
| ComplainingWitness | Text | 255 | yes |  |
| DOB | DateTime | 8 | yes |  |
| WkAddress | Text | 255 | yes |  |
| WkCity | Text | 255 | yes |  |
| WkState | Text | 255 | yes |  |
| WkZip | Text | 255 | yes |  |
| Pro Bono | Boolean | 1 | yes |  |
| HandlingAtty_Case | Text | 255 | yes |  |
| Action_Needed_on_Payment | Boolean | 1 | yes |  |
| SSN | Text | 255 | yes |  |
| Employer Name | Text | 255 | yes |  |
| Last Updated Contact Info | DateTime | 8 | yes |  |
| Ocounsel | Text | 255 | yes |  |
| Firm | Text | 255 | yes |  |
| OC Address | Text | 255 | yes |  |
| OC City | Text | 255 | yes |  |
| OC State | Text | 255 | yes |  |
| OC Zip | Text | 255 | yes |  |
| OC Phone | Text | 255 | yes |  |
| OC Email | Text | 255 | yes |  |
| OC Fax | Text | 255 | yes |  |
| Pro Bono PM | Text | 255 | yes |  |
| Pro Bono JRT | Text | 255 | yes |  |
| ContingencyFee | Boolean | 1 | yes |  |
| AuthorityToTalkTo | Memo | 0 | yes |  |
| Hourly | Boolean | 1 | yes |  |
| Contingency | Boolean | 1 | yes |  |
| Hybrid | Boolean | 1 | yes |  |
| Family-Law | Boolean | 1 | yes |  |
| Fixed | Boolean | 1 | yes |  |
| Scan | Boolean | 1 | yes |  |
| Scan Location | Memo | 0 | yes |  |
| ScanNotAvail | Boolean | 1 | yes |  |
| ParaLegal | Text | 255 | yes |  |
| Spanish | Boolean | 1 | yes |  |
| Offdate | DateTime | 8 | yes |  |
| CostHold | Currency | 8 | yes |  |
| CltNarrative | Memo | 0 | yes |  |
| ARTrustZero | Boolean | 1 | yes |  |
| F73 | Text | 255 | yes |  |
| F74 | Text | 255 | yes |  |
| F75 | Text | 255 | yes |  |
| F76 | Text | 255 | yes |  |
| PhName1 | Text | 255 | yes |  |
| PhName2 | Text | 255 | yes |  |
| LengthRes | Text | 255 | yes |  |
| LengthEmp | Text | 255 | yes |  |
| LegalStatus | Text | 255 | yes |  |
| CurrentBond | Text | 255 | yes |  |
| CrRecord | Memo | 0 | yes |  |
| TrustChronMemo | Memo | 0 | yes |  |
| Executor | Text | 255 | yes |  |
| RetainerReimb | Boolean | 1 | yes |  |
| RetReimbAmount | Currency | 8 | yes |  |
| Reviewable | Boolean | 1 | yes |  |
| ReviewReq | DateTime | 8 | yes |  |
| ReviewReceivedDate | DateTime | 8 | yes |  |
| ReviewReceived | Boolean | 1 | yes |  |
| Testimonial | Memo | 0 | yes |  |
| ReviewFollowUp | DateTime | 8 | yes |  |
| Stars | Long | 4 | yes |  |
| Review Source | Text | 255 | yes |  |
| Review Date | DateTime | 8 | yes |  |
| Title | Text | 255 | yes |  |
| OPartyLast | Text | 255 | yes |  |
| OPartyFirst | Text | 255 | yes |  |
| OPartyDOB | DateTime | 8 | yes |  |
| SSMA_TimeStamp | Binary | 8 | yes |  |
| SumOfDebit | Currency | 8 | yes |  |
| SumOfCredit | Currency | 8 | yes |  |
| Balance | Currency | 8 | yes |  |
| TrustAccountID | Long | 4 | yes |  |
| TDate | DateTime | 8 | yes |  |
| TMatter | Memo | 0 | yes |  |
| Debit | Currency | 8 | yes |  |
| Credit | Currency | 8 | yes |  |
| CheckNumber | Text | 255 | yes |  |
| Case No | Memo | 0 | yes |  |
| CheckCashed | Boolean | 1 | yes |  |
| DepCleared | Boolean | 1 | yes |  |
| Reconciled | Boolean | 1 | yes |  |
| OrderNr | Long | 4 | yes |  |
| Name | Memo | 0 | no |  |
| FileNumber | Memo | 0 | yes |  |

**Primary Key:** CaseID

**Linked Connect:** `ODBC;DRIVER=SQL Server;SERVER=awsql2022dev;APP=Microsoft Office;DATABASE=TateBywater`

**Indexes:**
- UniqueIndex [primary, unique]: CaseID

### vwTakeOff_A (linked)
| Column | Type | Size | Nullable | Default |
|--------|------|------|----------|---------|
| CaseID | Long | 4 | no |  |
| FileNumber | Memo | 0 | yes |  |
| Name | Memo | 0 | no |  |
| Orig_Atty | Text | 255 | yes |  |
| HandlingAtty_Case | Text | 255 | yes |  |
| Matter_type | Text | 255 | yes |  |
| Last_Name | Text | 255 | yes |  |
| SumOfAdvancedAR | Currency | 8 | yes |  |
| SumOfUnclearedDeposits | Currency | 8 | yes |  |
| Balance | Currency | 8 | yes |  |
| SumOfUncashedChecks | Currency | 8 | yes |  |
| SumOfTotal | Double | 8 | yes |  |
| IANumber | Text | 255 | yes |  |
| Bill_ID | Long | 4 | yes |  |
| CostHold | Currency | 8 | yes |  |
| SumOfCostBalance | Currency | 8 | yes |  |
| SumofPrepaid | Currency | 8 | yes |  |
| SumAdvLegal | Currency | 8 | yes |  |
| SumEarnedAdv | Currency | 8 | yes |  |
| SumCostReimb | Currency | 8 | yes |  |

**Primary Key:** CaseID

**Linked Connect:** `ODBC;DRIVER=SQL Server;SERVER=awsql2022dev;APP=Microsoft Office;DATABASE=TateBywater`

**Indexes:**
- UniqueIndex [primary, unique]: CaseID

### vwTakeOff_trust_account (linked)
| Column | Type | Size | Nullable | Default |
|--------|------|------|----------|---------|
| CaseID | Long | 4 | yes |  |
| SumOfDebit | Currency | 8 | yes |  |
| SumOfCredit | Currency | 8 | yes |  |
| Balance | Currency | 8 | yes |  |

**Primary Key:** CaseID

**Linked Connect:** `ODBC;DRIVER=SQL Server;SERVER=awsql2022dev;APP=Microsoft Office;DATABASE=TateBywater`

**Indexes:**
- UniqueIndex [primary, unique]: CaseID

### vwTakeOffStep2 (linked)
| Column | Type | Size | Nullable | Default |
|--------|------|------|----------|---------|
| FileNumber | Memo | 0 | yes |  |
| Name | Memo | 0 | no |  |
| CaseID | Long | 4 | no |  |
| Last_Name | Text | 255 | yes |  |
| First_Name | Text | 255 | yes |  |
| CaseOpenDate | DateTime | 8 | yes |  |
| Closed | Boolean | 1 | yes |  |
| Clsdate | DateTime | 8 | yes |  |
| Extended_Ledger | Text | 255 | yes |  |
| Case_Letter | Text | 255 | yes |  |
| yr | Text | 254 | yes |  |
| Number_ | Long | 4 | yes |  |
| Orig_Atty | Text | 255 | yes |  |
| Address | Text | 255 | yes |  |
| CourtCaseNo | Text | 255 | yes |  |
| City | Text | 255 | yes |  |
| FamilyLaw | Boolean | 1 | yes |  |
| State | Text | 255 | yes |  |
| Zip | Text | 255 | yes |  |
| Country | Text | 255 | yes |  |
| HmPhone | Text | 255 | yes |  |
| Action | Text | 255 | yes |  |
| OtherPhone | Text | 255 | yes |  |
| Fax | Text | 255 | yes |  |
| WkPhone | Text | 255 | yes |  |
| Comments | Memo | 0 | yes |  |
| Email | Text | 255 | yes |  |
| Referral | Text | 255 | yes |  |
| Individual Referrer | Text | 255 | yes |  |
| Retainer | Currency | 8 | yes |  |
| Matter_type | Text | 255 | yes |  |
| SOL | DateTime | 8 | yes |  |
| Court | Text | 255 | yes |  |
| CType | Text | 255 | yes |  |
| POfc | Text | 255 | yes |  |
| ComplainingWitness | Text | 255 | yes |  |
| DOB | DateTime | 8 | yes |  |
| WkAddress | Text | 255 | yes |  |
| WkCity | Text | 255 | yes |  |
| WkState | Text | 255 | yes |  |
| WkZip | Text | 255 | yes |  |
| Pro Bono | Boolean | 1 | yes |  |
| HandlingAtty_Case | Text | 255 | yes |  |
| Action_Needed_on_Payment | Boolean | 1 | yes |  |
| SSN | Text | 255 | yes |  |
| Employer Name | Text | 255 | yes |  |
| Last Updated Contact Info | DateTime | 8 | yes |  |
| Ocounsel | Text | 255 | yes |  |
| Firm | Text | 255 | yes |  |
| OC Address | Text | 255 | yes |  |
| OC City | Text | 255 | yes |  |
| OC State | Text | 255 | yes |  |
| OC Zip | Text | 255 | yes |  |
| OC Phone | Text | 255 | yes |  |
| OC Email | Text | 255 | yes |  |
| OC Fax | Text | 255 | yes |  |
| Pro Bono PM | Text | 255 | yes |  |
| Pro Bono JRT | Text | 255 | yes |  |
| ContingencyFee | Boolean | 1 | yes |  |
| AuthorityToTalkTo | Memo | 0 | yes |  |
| Hourly | Boolean | 1 | yes |  |
| Contingency | Boolean | 1 | yes |  |
| Hybrid | Boolean | 1 | yes |  |
| Family-Law | Boolean | 1 | yes |  |
| Fixed | Boolean | 1 | yes |  |
| Scan | Boolean | 1 | yes |  |
| Scan Location | Memo | 0 | yes |  |
| ScanNotAvail | Boolean | 1 | yes |  |
| ParaLegal | Text | 255 | yes |  |
| Spanish | Boolean | 1 | yes |  |
| Offdate | DateTime | 8 | yes |  |
| CostHold | Currency | 8 | yes |  |
| CltNarrative | Memo | 0 | yes |  |
| ARTrustZero | Boolean | 1 | yes |  |
| F73 | Text | 255 | yes |  |
| F74 | Text | 255 | yes |  |
| F75 | Text | 255 | yes |  |
| F76 | Text | 255 | yes |  |
| PhName1 | Text | 255 | yes |  |
| PhName2 | Text | 255 | yes |  |
| LengthRes | Text | 255 | yes |  |
| LengthEmp | Text | 255 | yes |  |
| LegalStatus | Text | 255 | yes |  |
| CurrentBond | Text | 255 | yes |  |
| CrRecord | Memo | 0 | yes |  |
| TrustChronMemo | Memo | 0 | yes |  |
| Executor | Text | 255 | yes |  |
| RetainerReimb | Boolean | 1 | yes |  |
| RetReimbAmount | Currency | 8 | yes |  |
| Reviewable | Boolean | 1 | yes |  |
| ReviewReq | DateTime | 8 | yes |  |
| ReviewReceivedDate | DateTime | 8 | yes |  |
| ReviewReceived | Boolean | 1 | yes |  |
| Testimonial | Memo | 0 | yes |  |
| ReviewFollowUp | DateTime | 8 | yes |  |
| Stars | Long | 4 | yes |  |
| Review Source | Text | 255 | yes |  |
| Review Date | DateTime | 8 | yes |  |
| Title | Text | 255 | yes |  |
| OPartyLast | Text | 255 | yes |  |
| OPartyFirst | Text | 255 | yes |  |
| OPartyDOB | DateTime | 8 | yes |  |
| TakeOffID | Long | 4 | yes |  |
| TakeOffMonthID | Long | 4 | yes |  |
| AvailBalance | Currency | 8 | yes |  |
| TotalUnCashedChks | Currency | 8 | yes |  |
| TotalUnclearedDeps | Currency | 8 | yes |  |
| TotalAdvancedAR | Currency | 8 | yes |  |
| EarlyEarned | Currency | 8 | yes |  |
| TOEarned | Currency | 8 | yes |  |
| TOAttBilled | Currency | 8 | yes |  |
| CostReimb | Currency | 8 | yes |  |
| CBHRev | Currency | 8 | yes |  |
| MKRev | Currency | 8 | yes |  |
| CBHCom | Currency | 8 | yes |  |
| MTRev | Currency | 8 | yes |  |
| MTCom | Currency | 8 | yes |  |
| KBCom | Currency | 8 | yes |  |
| MKCom | Currency | 8 | yes |  |
| TOEarnedTr | Boolean | 1 | yes |  |
| CostReimbTr | Boolean | 1 | yes |  |
| InsertedTrust | Boolean | 1 | yes |  |
| TotalHourlyOuts | Currency | 8 | yes |  |
| OpenTK | Text | 20 | yes |  |
| AdvCostBal | Currency | 8 | yes |  |
| AdvFeeBal | Currency | 8 | yes |  |
| CostHoldBal | Currency | 8 | yes |  |
| BRRev | Currency | 8 | yes |  |
| BRCom | Currency | 8 | yes |  |
| RLFCom | Currency | 8 | yes |  |
| AdvEarned | Currency | 8 | no |  |
| RemEarned | Currency | 8 | yes |  |

**Primary Key:** TakeOffID

**Linked Connect:** `ODBC;DRIVER=SQL Server;SERVER=awsql2022dev;APP=Microsoft Office;DATABASE=TateBywater`

**Indexes:**
- UniqueIndex [primary, unique]: TakeOffID

### vwTimeKeepingClosed (linked)
| Column | Type | Size | Nullable | Default |
|--------|------|------|----------|---------|
| Bill Sent | DateTime | 8 | yes |  |
| Bill Paid | DateTime | 8 | yes |  |
| Bill Closed | Boolean | 1 | yes |  |
| BilL Closed Date | DateTime | 8 | yes |  |
| Discount | Currency | 8 | yes |  |
| IANumber | Text | 255 | yes |  |
| FileNumber | Memo | 0 | yes |  |
| BalanceCalculated | Double | 8 | yes |  |
| CaseID | Long | 4 | no |  |
| Last_Name | Text | 255 | yes |  |
| First_Name | Text | 255 | yes |  |
| CaseOpenDate | DateTime | 8 | yes |  |
| Closed | Boolean | 1 | yes |  |
| Clsdate | DateTime | 8 | yes |  |
| Extended_Ledger | Text | 255 | yes |  |
| Case_Letter | Text | 255 | yes |  |
| yr | Text | 254 | yes |  |
| Number_ | Long | 4 | yes |  |
| Orig_Atty | Text | 255 | yes |  |
| Address | Text | 255 | yes |  |
| CourtCaseNo | Text | 255 | yes |  |
| City | Text | 255 | yes |  |
| FamilyLaw | Boolean | 1 | yes |  |
| State | Text | 255 | yes |  |
| Zip | Text | 255 | yes |  |
| Country | Text | 255 | yes |  |
| HmPhone | Text | 255 | yes |  |
| Action | Text | 255 | yes |  |
| OtherPhone | Text | 255 | yes |  |
| Fax | Text | 255 | yes |  |
| WkPhone | Text | 255 | yes |  |
| Comments | Memo | 0 | yes |  |
| Email | Text | 255 | yes |  |
| Referral | Text | 255 | yes |  |
| Individual Referrer | Text | 255 | yes |  |
| Retainer | Currency | 8 | yes |  |
| Matter_type | Text | 255 | yes |  |
| SOL | DateTime | 8 | yes |  |
| Court | Text | 255 | yes |  |
| CType | Text | 255 | yes |  |
| POfc | Text | 255 | yes |  |
| ComplainingWitness | Text | 255 | yes |  |
| DOB | DateTime | 8 | yes |  |
| WkAddress | Text | 255 | yes |  |
| WkCity | Text | 255 | yes |  |
| WkState | Text | 255 | yes |  |
| WkZip | Text | 255 | yes |  |
| Pro Bono | Boolean | 1 | yes |  |
| HandlingAtty_Case | Text | 255 | yes |  |
| Action_Needed_on_Payment | Boolean | 1 | yes |  |
| SSN | Text | 255 | yes |  |
| Employer Name | Text | 255 | yes |  |
| Last Updated Contact Info | DateTime | 8 | yes |  |
| Ocounsel | Text | 255 | yes |  |
| Firm | Text | 255 | yes |  |
| OC Address | Text | 255 | yes |  |
| OC City | Text | 255 | yes |  |
| OC State | Text | 255 | yes |  |
| OC Zip | Text | 255 | yes |  |
| OC Phone | Text | 255 | yes |  |
| OC Email | Text | 255 | yes |  |
| OC Fax | Text | 255 | yes |  |
| Pro Bono PM | Text | 255 | yes |  |
| Pro Bono JRT | Text | 255 | yes |  |
| ContingencyFee | Boolean | 1 | yes |  |
| AuthorityToTalkTo | Memo | 0 | yes |  |
| Hourly | Boolean | 1 | yes |  |
| Contingency | Boolean | 1 | yes |  |
| Hybrid | Boolean | 1 | yes |  |
| Family-Law | Boolean | 1 | yes |  |
| Fixed | Boolean | 1 | yes |  |
| Scan | Boolean | 1 | yes |  |
| Scan Location | Memo | 0 | yes |  |
| ScanNotAvail | Boolean | 1 | yes |  |
| ParaLegal | Text | 255 | yes |  |
| Spanish | Boolean | 1 | yes |  |
| Offdate | DateTime | 8 | yes |  |
| CostHold | Currency | 8 | yes |  |
| CltNarrative | Memo | 0 | yes |  |
| ARTrustZero | Boolean | 1 | yes |  |
| F73 | Text | 255 | yes |  |
| F74 | Text | 255 | yes |  |
| F75 | Text | 255 | yes |  |
| F76 | Text | 255 | yes |  |
| PhName1 | Text | 255 | yes |  |
| PhName2 | Text | 255 | yes |  |
| LengthRes | Text | 255 | yes |  |
| LengthEmp | Text | 255 | yes |  |
| LegalStatus | Text | 255 | yes |  |
| CurrentBond | Text | 255 | yes |  |
| CrRecord | Memo | 0 | yes |  |
| TrustChronMemo | Memo | 0 | yes |  |
| Executor | Text | 255 | yes |  |
| RetainerReimb | Boolean | 1 | yes |  |
| RetReimbAmount | Currency | 8 | yes |  |
| Reviewable | Boolean | 1 | yes |  |
| ReviewReq | DateTime | 8 | yes |  |
| ReviewReceivedDate | DateTime | 8 | yes |  |
| ReviewReceived | Boolean | 1 | yes |  |
| Testimonial | Memo | 0 | yes |  |
| ReviewFollowUp | DateTime | 8 | yes |  |
| Stars | Long | 4 | yes |  |
| Review Source | Text | 255 | yes |  |
| Review Date | DateTime | 8 | yes |  |
| Title | Text | 255 | yes |  |
| OPartyLast | Text | 255 | yes |  |
| OPartyFirst | Text | 255 | yes |  |
| OPartyDOB | DateTime | 8 | yes |  |
| SSMA_TimeStamp | Binary | 8 | no |  |
| Bill Open | DateTime | 8 | yes |  |
| Name | Memo | 0 | no |  |
| Bill_ID | Long | 4 | no |  |
| TrustatClose | Currency | 8 | yes |  |
| StatementLessTrust | Boolean | 1 | yes |  |
| InvoiceExceedsTrust | Boolean | 1 | yes |  |
| InvoiceTotalAdvance | Boolean | 1 | yes |  |
| InvoiceNoAdvance | Boolean | 1 | yes |  |

**Primary Key:** CaseID

**Linked Connect:** `ODBC;DRIVER=SQL Server;SERVER=awsql2022dev;APP=Microsoft Office;DATABASE=TateBywater`

**Indexes:**
- UniqueIndex [primary, unique]: CaseID

### vwTimeKeepingOpen (linked)
| Column | Type | Size | Nullable | Default |
|--------|------|------|----------|---------|
| Bill Closed | Boolean | 1 | yes |  |
| Bill_ID | Long | 4 | no |  |
| IANumber | Text | 255 | yes |  |
| FileNumber | Memo | 0 | yes |  |
| BalanceCalculated | Double | 8 | yes |  |
| CaseID | Long | 4 | no |  |
| Last_Name | Text | 255 | yes |  |
| First_Name | Text | 255 | yes |  |
| CaseOpenDate | DateTime | 8 | yes |  |
| Closed | Boolean | 1 | yes |  |
| Clsdate | DateTime | 8 | yes |  |
| Extended_Ledger | Text | 255 | yes |  |
| Case_Letter | Text | 255 | yes |  |
| yr | Text | 254 | yes |  |
| Number_ | Long | 4 | yes |  |
| Orig_Atty | Text | 255 | yes |  |
| Address | Text | 255 | yes |  |
| CourtCaseNo | Text | 255 | yes |  |
| City | Text | 255 | yes |  |
| FamilyLaw | Boolean | 1 | yes |  |
| State | Text | 255 | yes |  |
| Zip | Text | 255 | yes |  |
| Country | Text | 255 | yes |  |
| HmPhone | Text | 255 | yes |  |
| Action | Text | 255 | yes |  |
| OtherPhone | Text | 255 | yes |  |
| Fax | Text | 255 | yes |  |
| WkPhone | Text | 255 | yes |  |
| Comments | Memo | 0 | yes |  |
| Email | Text | 255 | yes |  |
| Referral | Text | 255 | yes |  |
| Individual Referrer | Text | 255 | yes |  |
| Retainer | Currency | 8 | yes |  |
| Matter_type | Text | 255 | yes |  |
| SOL | DateTime | 8 | yes |  |
| Court | Text | 255 | yes |  |
| CType | Text | 255 | yes |  |
| POfc | Text | 255 | yes |  |
| ComplainingWitness | Text | 255 | yes |  |
| DOB | DateTime | 8 | yes |  |
| WkAddress | Text | 255 | yes |  |
| WkCity | Text | 255 | yes |  |
| WkState | Text | 255 | yes |  |
| WkZip | Text | 255 | yes |  |
| Pro Bono | Boolean | 1 | yes |  |
| HandlingAtty_Case | Text | 255 | yes |  |
| Action_Needed_on_Payment | Boolean | 1 | yes |  |
| SSN | Text | 255 | yes |  |
| Employer Name | Text | 255 | yes |  |
| Last Updated Contact Info | DateTime | 8 | yes |  |
| Ocounsel | Text | 255 | yes |  |
| Firm | Text | 255 | yes |  |
| OC Address | Text | 255 | yes |  |
| OC City | Text | 255 | yes |  |
| OC State | Text | 255 | yes |  |
| OC Zip | Text | 255 | yes |  |
| OC Phone | Text | 255 | yes |  |
| OC Email | Text | 255 | yes |  |
| OC Fax | Text | 255 | yes |  |
| Pro Bono PM | Text | 255 | yes |  |
| Pro Bono JRT | Text | 255 | yes |  |
| ContingencyFee | Boolean | 1 | yes |  |
| AuthorityToTalkTo | Memo | 0 | yes |  |
| Hourly | Boolean | 1 | yes |  |
| Contingency | Boolean | 1 | yes |  |
| Hybrid | Boolean | 1 | yes |  |
| Family-Law | Boolean | 1 | yes |  |
| Fixed | Boolean | 1 | yes |  |
| Scan | Boolean | 1 | yes |  |
| Scan Location | Memo | 0 | yes |  |
| ScanNotAvail | Boolean | 1 | yes |  |
| ParaLegal | Text | 255 | yes |  |
| Spanish | Boolean | 1 | yes |  |
| Offdate | DateTime | 8 | yes |  |
| CostHold | Currency | 8 | yes |  |
| CltNarrative | Memo | 0 | yes |  |
| ARTrustZero | Boolean | 1 | yes |  |
| F73 | Text | 255 | yes |  |
| F74 | Text | 255 | yes |  |
| F75 | Text | 255 | yes |  |
| F76 | Text | 255 | yes |  |
| PhName1 | Text | 255 | yes |  |
| PhName2 | Text | 255 | yes |  |
| LengthRes | Text | 255 | yes |  |
| LengthEmp | Text | 255 | yes |  |
| LegalStatus | Text | 255 | yes |  |
| CurrentBond | Text | 255 | yes |  |
| CrRecord | Memo | 0 | yes |  |
| TrustChronMemo | Memo | 0 | yes |  |
| Executor | Text | 255 | yes |  |
| RetainerReimb | Boolean | 1 | yes |  |
| RetReimbAmount | Currency | 8 | yes |  |
| Reviewable | Boolean | 1 | yes |  |
| ReviewReq | DateTime | 8 | yes |  |
| ReviewReceivedDate | DateTime | 8 | yes |  |
| ReviewReceived | Boolean | 1 | yes |  |
| Testimonial | Memo | 0 | yes |  |
| ReviewFollowUp | DateTime | 8 | yes |  |
| Stars | Long | 4 | yes |  |
| Review Source | Text | 255 | yes |  |
| Review Date | DateTime | 8 | yes |  |
| Title | Text | 255 | yes |  |
| OPartyLast | Text | 255 | yes |  |
| OPartyFirst | Text | 255 | yes |  |
| OPartyDOB | DateTime | 8 | yes |  |
| SSMA_TimeStamp | Binary | 8 | no |  |
| Bill Open | DateTime | 8 | yes |  |
| Name | Memo | 0 | no |  |

**Primary Key:** CaseID

**Linked Connect:** `ODBC;DRIVER=SQL Server;SERVER=awsql2022dev;APP=Microsoft Office;DATABASE=TateBywater`

**Indexes:**
- UniqueIndex [primary, unique]: CaseID

### vwTimeTableDetail (linked)
| Column | Type | Size | Nullable | Default |
|--------|------|------|----------|---------|
| Time_ID | Long | 4 | no |  |
| Bill_ID | Long | 4 | yes |  |
| Tdate | DateTime | 8 | yes |  |
| Description | Memo | 0 | yes |  |
| Tatty | Text | 255 | yes |  |
| Rate | Currency | 8 | yes |  |
| Time_ | Double | 8 | yes |  |
| Amount | Double | 8 | yes |  |

**Primary Key:** Time_ID

**Linked Connect:** `ODBC;DRIVER=SQL Server;SERVER=awsql2022dev;APP=Microsoft Office;DATABASE=TateBywater`

**Indexes:**
- UniqueIndex [primary, unique]: Time_ID

### vwTKClose_A (linked)
| Column | Type | Size | Nullable | Default |
|--------|------|------|----------|---------|
| CaseID | Long | 4 | no |  |
| FileNumber | Memo | 0 | yes |  |
| Name | Memo | 0 | no |  |
| Orig_Atty | Text | 255 | yes |  |
| Case_Letter | Text | 255 | yes |  |
| HandlingAtty_Case | Text | 255 | yes |  |
| SumOfAdvancedAR | Currency | 8 | yes |  |
| CostHold | Currency | 8 | yes |  |
| SumOfUnclearedDeposits | Currency | 8 | yes |  |
| Balance | Currency | 8 | yes |  |
| SumOfUncashedChecks | Currency | 8 | yes |  |
| AvailBalance | Currency | 8 | yes |  |
| BankBalance | Currency | 8 | yes |  |
| SumOfTotal | Double | 8 | yes |  |
| IANumber | Text | 255 | yes |  |
| Bill_ID | Long | 4 | yes |  |
| Retainer | Currency | 8 | yes |  |
| RetainerReimb | Boolean | 1 | yes |  |
| RetReimbAmount | Currency | 8 | yes |  |
| MaxOfMatterID | Long | 4 | yes |  |
| AROutstandingBalance | Currency | 8 | yes |  |
| AdvanceCostBalance | Currency | 8 | yes |  |
| AdvLegalFeeBalance | Currency | 8 | yes |  |
| CostExpendedBalance | Currency | 8 | yes |  |
| Fixed | Boolean | 1 | yes |  |

**Primary Key:** CaseID

**Linked Connect:** `ODBC;DRIVER=SQL Server;SERVER=awsql2022dev;APP=Microsoft Office;DATABASE=TateBywater`

**Indexes:**
- UniqueIndex [primary, unique]: CaseID

### vwTrustAccount (linked)
| Column | Type | Size | Nullable | Default |
|--------|------|------|----------|---------|
| CaseID | Long | 4 | yes |  |
| SumOfDebit | Currency | 8 | yes |  |
| SumOfCredit | Currency | 8 | yes |  |
| Balance | Currency | 8 | yes |  |
| OrderNr | Long | 4 | yes |  |
| TrustAccountID | Long | 4 | no |  |

**Primary Key:** TrustAccountID

**Linked Connect:** `ODBC;DRIVER=SQL Server;SERVER=awsql2022dev;APP=Microsoft Office;DATABASE=TateBywater`

**Indexes:**
- UniqueIndex [primary, unique]: TrustAccountID

### vwTrustAccountBalanceTotals (linked)
| Column | Type | Size | Nullable | Default |
|--------|------|------|----------|---------|
| CaseID | Long | 4 | yes |  |
| SumOfBalance | Currency | 8 | yes |  |

**Primary Key:** CaseID

**Linked Connect:** `ODBC;DRIVER=SQL Server;SERVER=awsql2022dev;APP=Microsoft Office;DATABASE=TateBywater`

**Indexes:**
- UniqueIndex [primary, unique]: CaseID

### vwTrustAccountTable (linked)
| Column | Type | Size | Nullable | Default |
|--------|------|------|----------|---------|
| TrustAccountID | Long | 4 | no |  |
| CaseID | Long | 4 | yes |  |
| TDate | DateTime | 8 | yes |  |
| TMatter | Memo | 0 | yes |  |
| Debit | Currency | 8 | yes |  |
| Credit | Currency | 8 | yes |  |
| CheckCashed | Boolean | 1 | yes |  |
| CheckNumber | Text | 255 | yes |  |
| DepCleared | Boolean | 1 | yes |  |
| Reconciled | Boolean | 1 | yes |  |
| OrderNr | Long | 4 | yes |  |
| CostReimb | Boolean | 1 | yes |  |
| AdvFee | Boolean | 1 | yes |  |
| SumOfDebit | Currency | 8 | yes |  |
| SumOfCredit | Currency | 8 | yes |  |
| Balance | Currency | 8 | yes |  |

**Primary Key:** TrustAccountID

**Linked Connect:** `ODBC;DRIVER=SQL Server;SERVER=awsql2022dev;APP=Microsoft Office;DATABASE=TateBywater`

**Indexes:**
- UniqueIndex [primary, unique]: TrustAccountID

### vwTrustCostsExpendedTotals (linked)
| Column | Type | Size | Nullable | Default |
|--------|------|------|----------|---------|
| SumOfCostBalance | Currency | 8 | yes |  |
| CaseID | Long | 4 | no |  |

**Primary Key:** CaseID

**Linked Connect:** `ODBC;DRIVER=SQL Server;SERVER=awsql2022dev;APP=Microsoft Office;DATABASE=TateBywater`

**Indexes:**
- UniqueIndex [primary, unique]: CaseID

### vwTrustReportRPT (linked)
| Column | Type | Size | Nullable | Default |
|--------|------|------|----------|---------|
| CaseID | Long | 4 | no |  |
| Case_Letter | Text | 255 | yes |  |
| yr | Text | 254 | yes |  |
| Number_ | Long | 4 | yes |  |
| Orig_Atty | Text | 255 | yes |  |
| Matter_type | Text | 255 | yes |  |
| CaseOpenDate | DateTime | 8 | yes |  |
| CheckCashed | Boolean | 1 | yes |  |
| CaseNo | Memo | 0 | yes |  |
| Last_Name | Text | 255 | yes |  |
| First_Name | Text | 255 | yes |  |
| TrustAccountID | Long | 4 | yes |  |

**Primary Key:** CaseID

**Linked Connect:** `ODBC;DRIVER=SQL Server;SERVER=awsql2022dev;APP=Microsoft Office;DATABASE=TateBywater`

**Indexes:**
- UniqueIndex [primary, unique]: CaseID

### vwTrustReportRPT1 (linked)
| Column | Type | Size | Nullable | Default |
|--------|------|------|----------|---------|
| CaseID | Long | 4 | no |  |
| Case_Letter | Text | 255 | yes |  |
| yr | Text | 254 | yes |  |
| Number_ | Long | 4 | yes |  |
| Orig_Atty | Text | 255 | yes |  |
| Matter_type | Text | 255 | yes |  |
| CaseOpenDate | DateTime | 8 | yes |  |
| SumOfBalance | Long | 4 | no |  |
| CheckCashed | Boolean | 1 | yes |  |
| CaseNo | Memo | 0 | yes |  |
| Last_Name | Text | 255 | yes |  |
| First_Name | Text | 255 | yes |  |

**Primary Key:** CaseID

**Linked Connect:** `ODBC;DRIVER=SQL Server;SERVER=awsql2022dev;APP=Microsoft Office;DATABASE=TateBywater`

**Indexes:**
- UniqueIndex [primary, unique]: CaseID

### vwTrustTotalEarnedSum (linked)
| Column | Type | Size | Nullable | Default |
|--------|------|------|----------|---------|
| SumOfSumOfCredit | Currency | 8 | yes |  |
| CaseID | Long | 4 | no |  |

**Primary Key:** CaseID

**Linked Connect:** `ODBC;DRIVER=SQL Server;SERVER=awsql2022dev;APP=Microsoft Office;DATABASE=TateBywater`

**Indexes:**
- UniqueIndex [primary, unique]: CaseID

### vwTTAmountHours_SUM_byAtty_TotalCaseID (linked)
| Column | Type | Size | Nullable | Default |
|--------|------|------|----------|---------|
| SumOfSumOfTime_ | Double | 8 | yes |  |
| Tatty | Text | 255 | yes |  |
| CaseID | Long | 4 | yes |  |

**Primary Key:** CaseID

**Linked Connect:** `ODBC;DRIVER=SQL Server;SERVER=awsql2022dev;APP=Microsoft Office;DATABASE=TateBywater`

**Indexes:**
- UniqueIndex [primary, unique]: CaseID

### vwTTAmountHours_TotalCaseID (linked)
| Column | Type | Size | Nullable | Default |
|--------|------|------|----------|---------|
| SumOfSumOfTime_ | Double | 8 | yes |  |
| CaseID | Long | 4 | yes |  |

**Primary Key:** CaseID

**Linked Connect:** `ODBC;DRIVER=SQL Server;SERVER=awsql2022dev;APP=Microsoft Office;DATABASE=TateBywater`

**Indexes:**
- UniqueIndex [primary, unique]: CaseID

### z_PCADataSources
| Column | Type | Size | Nullable | Default |
|--------|------|------|----------|---------|
| PCADataSourceName | Text | 200 | no | "BaseTables" |
| ApplicationStatus | Text | 50 | no |  |
| DataSourceType | Text | 50 | yes | "" |
| DataSourceConnectString | Text | 255 | yes | "" |
| TablePrefix | Text | 50 | yes | "" |
| TableSuffix | Text | 50 | yes | "" |
| VisibleInSetupForm | Integer | 2 | yes | 0 |
| MustBeConnectedForAppToRun | Integer | 2 | yes | 0 |
| SQLServerName | Text | 255 | yes |  |

**Primary Key:** PCADataSourceName, ApplicationStatus

**Indexes:**
- PrimaryKey [primary, unique]: PCADataSourceName, ApplicationStatus

### z_PCADataSources_TableList
| Column | Type | Size | Nullable | Default |
|--------|------|------|----------|---------|
| PCADataSourceName | Text | 50 | no | "TateBywater" |
| ConnectAs | Text | 50 | no | "" |
| ForeignTableName | Text | 50 | yes | "" |
| IsView | Boolean | 1 | yes | No |
| UniqueID | Text | 255 | yes |  |

**Primary Key:** PCADataSourceName, ConnectAs

**Indexes:**
- PrimaryKey [primary, unique]: PCADataSourceName, ConnectAs
- UniqueID: UniqueID

### z_PCASettings
| Column | Type | Size | Nullable | Default |
|--------|------|------|----------|---------|
| INISection | Text | 50 | no |  |
| INIKey | Text | 50 | no | "" |
| INIDescription | Text | 50 | yes |  |

**Primary Key:** INISection, INIKey

**Indexes:**
- Key: INIKey
- PrimaryKey [primary, unique]: INISection, INIKey

## Relationships
No relationships extracted.

## Queries
### Query1
```sql
SELECT tblCase.CaseID, [Trust Account].TDate, [Trust Account].TMatter, [Trust Account].Credit, [Trust Account].OrderNr

FROM tblCase INNER JOIN [Trust Account] ON tblCase.CaseID = [Trust Account].CaseID

WHERE ((([Trust Account].TMatter) Not Like "*Earned*" And ([Trust Account].TMatter) Not Like "*Reimb*" And ([Trust Account].TMatter) Not Like "*Refund*") AND (([Trust Account].Credit)>0))

ORDER BY [Trust Account].OrderNr;
```
### Sele
```sql
SELECT tblDropD.CodeVal, tblDropD.FieldName, tblDropD.SortOrder

FROM tblDropD

WHERE (((tblDropD.FieldName)="TimeTableAtty"))

ORDER BY tblDropD.SortOrder;
```
### qryARCredits
```sql
SELECT [Matter and AR].CaseID, [Matter and AR].Payment, [Matter and AR].Pay_Outlay

FROM [Matter and AR]

WHERE ((([Matter and AR].Pay_Outlay) Like "*credit*"));
```
### qryARCreditsSum
```sql
SELECT vwARCreditsSum.CaseID, vwARCreditsSum.SumOfPayment

FROM vwARCreditsSum;
```
### qryARCreditsSum_OLD
```sql
SELECT qryARCredits.CaseID, Sum(qryARCredits.Payment) AS SumOfPayment

FROM qryARCredits

GROUP BY qryARCredits.CaseID;
```
### qryActionNeededAll
```sql
SELECT tblCase.CaseID, tblCase.Orig_Atty, tblCase.HandlingAtty_Case, tblCase.ParaLegal, tblCase.Matter_type, [case_Letter] & [yr] & "-" & [Number_] & "-" & [Orig_Atty] AS [Case No], [Last_Name] & ", " & [First_Name] AS Name, TblActionNeeded.ActionNeededDet, TblActionNeeded.ActionComp, tblCase.Action_Needed_on_Payment, tblCase.SOL, tblDropD.CodeVal, TblActionNeeded.DateComp, TblActionNeeded.ActPerson, TblActionNeeded.DateComp1, TblActionNeeded.StartDate, tblCase.Closed, TblActionNeeded.ActionNeed...
```
### qryActionNeededAll2
```sql
SELECT tblCase.CaseID, tblCase.Orig_Atty, tblCase.HandlingAtty_Case, tblCase.ParaLegal, tblCase.Matter_type, [case_Letter] & [yr] & "-" & [Number_] & "-" & [Orig_Atty] AS [Case No], [Last_Name] & ", " & [First_Name] AS Name, TblActionNeeded.ActionNeededDet, TblActionNeeded.ActionComp, tblCase.Action_Needed_on_Payment

FROM tblCase LEFT JOIN TblActionNeeded ON tblCase.CaseID = TblActionNeeded.CaseID

WHERE (((TblActionNeeded.ActionComp)=No));
```
### qryActionNeededAll3
```sql
SELECT tblCase.CaseID, tblCase.Orig_Atty, tblCase.HandlingAtty_Case, tblCase.ParaLegal, tblCase.Matter_type, [case_Letter] & [yr] & "-" & [Number_] & "-" & [Orig_Atty] AS [Case No], [Last_Name] & ", " & [First_Name] AS Name, TblActionNeeded.ActionNeededDet, TblActionNeeded.ActionComp, tblCase.Action_Needed_on_Payment

FROM tblCase LEFT JOIN TblActionNeeded ON tblCase.CaseID = TblActionNeeded.CaseID

WHERE (((TblActionNeeded.ActionComp)=No));
```
### qryActionNeededAllNEW
```sql
SELECT tblCase.CaseID, tblCase.Last_Name, tblCase.First_Name, tblCase.Case_Letter, tblCase.yr, tblCase.Number_, [case_Letter] & [yr] & "-" & [Number_] & "-" & [Orig_Atty] AS [Case No], [Last_Name] & ", " & [First_Name] AS Name, TblActionNeeded.ActionNeededDet, TblActionNeeded.ActionComp, tblCase.Orig_Atty, tblCase.HandlingAtty_Case, tblCase.ParaLegal, tblCase.Matter_type

FROM tblCase INNER JOIN TblActionNeeded ON tblCase.CaseID = TblActionNeeded.CaseID

WHERE (((TblActionNeeded.ActionComp)=No))...
```
### qryAdvLegalFees
```sql
SELECT [Matter and AR].CaseID, [Matter and AR].Charge, [Matter and AR].AdvancedLegal

FROM [Matter and AR]

WHERE ((([Matter and AR].AdvancedLegal)=True));
```
### qryAdvLegalFeesSum
```sql
SELECT vwAdvLegalFeesSum.CaseID, vwAdvLegalFeesSum.SumOfCharge, vwAdvLegalFeesSum.AdvancedLegal

FROM vwAdvLegalFeesSum;
```
### qryAdvLegalFeesSum_OLD
```sql
SELECT qryAdvLegalFees.CaseID, Sum(qryAdvLegalFees.Charge) AS SumOfCharge, qryAdvLegalFees.AdvancedLegal

FROM qryAdvLegalFees

GROUP BY qryAdvLegalFees.CaseID, qryAdvLegalFees.AdvancedLegal;
```
### qryAttyTrustAcctsTOff
```sql
SELECT Replace([Case_Letter] & [yr] & "-" & [Number_] & "-" & [Orig_Atty],"__","_") AS FileNumber, [Last_Name] & ", " & [First_Name] AS Name, tblCase.*, tblTakeOff.*

FROM tblCase INNER JOIN tblTakeOff ON tblCase.CaseID = tblTakeOff.CaseID;
```
### qryBillList
```sql
SELECT [TB Time Keeping].Bill_ID, [TB Time Keeping].IANumber, [TB Time Keeping].[BilL Closed Date], TblCase.CaseID

FROM TblCase INNER JOIN [TB Time Keeping] ON TblCase.CaseID = [TB Time Keeping].CaseID

ORDER BY [TB Time Keeping].Bill_ID, [TB Time Keeping].IANumber, TblCase.CaseID;
```
### qryBillingTotals
```sql
SELECT tblTimeTableDetail.Bill_ID, Sum(Nz([Time_],0)*Nz([Rate],0)) AS Amount

FROM tblTimeTableDetail

GROUP BY tblTimeTableDetail.Bill_ID;
```
### qryBillingTracker
```sql
SELECT tblTimeTableDetail.Tdate, tblTimeTableDetail.Tatty, tblTimeTableDetail.Bill_ID, tblTimeTableDetail.Time_

FROM tblTimeTableDetail

ORDER BY tblTimeTableDetail.Tdate DESC;
```
### qryBillingTracker2
```sql
SELECT vwBillingTracker2.Bill_ID, vwBillingTracker2.Tdate, vwBillingTracker2.Tatty, vwBillingTracker2.Rate, vwBillingTracker2.Time_, vwBillingTracker2.Billed, vwBillingTracker2.CaseID, vwBillingTracker2.Last_Name, vwBillingTracker2.First_Name, vwBillingTracker2.Case_Letter, vwBillingTracker2.yr, vwBillingTracker2.Number_, vwBillingTracker2.Orig_Atty, vwBillingTracker2.Name, vwBillingTracker2.FileNumber, vwBillingTracker2.Time_ID, *

FROM vwBillingTracker2

ORDER BY vwBillingTracker2.Tdate DESC;
```
### qryBillingTracker2_OLD
```sql
SELECT tblTimeTableDetail.Bill_ID, tblTimeTableDetail.Tdate, tblTimeTableDetail.Tatty, tblTimeTableDetail.Rate, tblTimeTableDetail.Time_, [Rate]*[Time_] AS Billed, [TB Time Keeping].CaseID, tblCase.Last_Name, tblCase.First_Name, tblCase.Case_Letter, tblCase.yr, tblCase.Number_, tblCase.Orig_Atty, [Last_Name] & ", " & [First_Name] AS Name, Replace([Case_Letter] & [yr] & "-" & [Number_] & "-" & [Orig_Atty],"__","_") AS FileNumber

FROM tblCase INNER JOIN ([TB Time Keeping] INNER JOIN tblTimeTableD...
```
### qryCalendarCheck
```sql
SELECT tblCase.CaseID, tblCase.Orig_Atty, tblCase.HandlingAtty_Case, tblCase.ParaLegal, tblCase.Matter_type, tblHearingDate.Hearing_Date, tblHearingDate.HearingType, [case_Letter] & [yr] & "-" & [Number_] & "-" & [Orig_Atty] AS [Case No], tblCase.Last_Name, tblCase.First_Name, tblCase.Case_Letter, tblCase.yr, tblCase.Number_, [Last_Name] & ", " & [First_Name] AS Name, tblHearingDate.HearingTime, tblHearingDate.HrgResult, tblCase.Closed

FROM tblCase INNER JOIN tblHearingDate ON tblCase.CaseID = ...
```
### qryCaseIDclientsAll
```sql
SELECT TblCase.CaseID, [last_name] & ", " & [first_name] AS Name

FROM TblCase

GROUP BY TblCase.CaseID, [last_name] & ", " & [first_name], TblCase.Last_Name

HAVING (((TblCase.Last_Name) Is Not Null))

ORDER BY [last_name] & ", " & [first_name];
```
### qryCaseIDclientsClosed
```sql
SELECT TblCase.CaseID, [last_name] & ", " & [first_name] AS Name, TblCase.Closed

FROM TblCase

GROUP BY TblCase.CaseID, [last_name] & ", " & [first_name], TblCase.Closed, TblCase.Last_Name

HAVING (((TblCase.Closed)=Yes) AND ((TblCase.Last_Name) Is Not Null))

ORDER BY [last_name] & ", " & [first_name];
```
### qryCaseIDclientsclosednotscanned
```sql
SELECT TblCase.CaseID, [last_name] & ", " & [first_name] AS Name, TblCase.Closed, TblCase.Scan, TblCase.[Scan Location], TblCase.ScanNotAvail

FROM TblCase

GROUP BY TblCase.CaseID, [last_name] & ", " & [first_name], TblCase.Closed, TblCase.Scan, TblCase.[Scan Location], TblCase.ScanNotAvail, TblCase.Last_Name

HAVING (((TblCase.Closed)=Yes) AND ((TblCase.Scan)=No) AND ((TblCase.ScanNotAvail)=No) AND ((TblCase.Last_Name) Is Not Null)) OR (((TblCase.[Scan Location])="No"))

ORDER BY [last_name] &...
```
### qryCaseList
```sql
SELECT tblCase.CaseID, tblCase.Matter_type, tblCase.Court, Replace([case_letter] & "_" & [yr] & "_" & [Number_] & "_" & [Orig_Atty],"__","_") AS CaseNo, tblCase.CourtCaseNo, tblCase.CaseOpenDate

FROM tblCase;
```
### qryCaseListAll
```sql
SELECT TblCase.CaseID, TblCase.CaseOpenDate, [Last_Name] & ", " & [First_Name] AS ClientName, TblCase.Case_Letter, TblCase.yr, TblCase.Number_, TblCase.Orig_Atty, TblCase.Extended_Ledger, TblCase.Court, TblCase.Matter_type, Replace([Case_Letter] & [yr] & "-" & [Number_] & "-" & [Orig_Atty],"__","_") AS FileNumber, TblCase.[Scan Location], TblCase.HandlingAtty_Case, TblCase.Closed, tblDropD.CodeVal, TblCase.ParaLegal

FROM TblCase LEFT JOIN tblDropD ON TblCase.Case_Letter = tblDropD.Code

ORDER B...
```
### qryCaseListClosed
```sql
SELECT TblCase.CaseID, TblCase.CaseOpenDate, [Last_Name] & ", " & [First_Name] AS ClientName, TblCase.Case_Letter, TblCase.yr, TblCase.Number_, TblCase.Orig_Atty, TblCase.Extended_Ledger, TblCase.Court, TblCase.Matter_type, Replace([Case_Letter] & [yr] & "-" & [Number_] & "-" & [Orig_Atty],"__","_") AS FileNumber, TblCase.[Scan Location], TblCase.HandlingAtty_Case, TblCase.Closed, tblDropD.CodeVal, TblCase.Clsdate, TblCase.ParaLegal

FROM TblCase LEFT JOIN tblDropD ON TblCase.Case_Letter = tblDr...
```
### qryCaseListOpen
```sql
SELECT vwCaseListOpen.CaseID, vwCaseListOpen.CaseOpenDate, vwCaseListOpen.ClientName, vwCaseListOpen.Case_Letter, vwCaseListOpen.yr, vwCaseListOpen.Number_, vwCaseListOpen.Orig_Atty, vwCaseListOpen.Extended_Ledger, vwCaseListOpen.Court, vwCaseListOpen.Matter_type, vwCaseListOpen.FileNumber, vwCaseListOpen.[Scan Location], vwCaseListOpen.HandlingAtty_Case, vwCaseListOpen.Closed, vwCaseListOpen.CodeVal, vwCaseListOpen.ParaLegal, vwCaseListOpen.PIStatus, vwCaseListOpen.Retainer

FROM vwCaseListOpen...
```
### qryCaseListOpen_OLD
```sql
SELECT TblCase.CaseID, TblCase.CaseOpenDate, [Last_Name] & ", " & [First_Name] AS ClientName, TblCase.Case_Letter, TblCase.yr, TblCase.Number_, TblCase.Orig_Atty, TblCase.Extended_Ledger, TblCase.Court, TblCase.Matter_type, Replace([Case_Letter] & [yr] & "-" & [Number_] & "-" & [Orig_Atty],"__","_") AS FileNumber, TblCase.[Scan Location], TblCase.HandlingAtty_Case, TblCase.Closed, tblDropD.CodeVal, TblCase.ParaLegal, [Personal Injury].PIStatus

FROM (TblCase LEFT JOIN tblDropD ON TblCase.Case_Le...
```
### qryCaseSourcesRPT
```sql
SELECT vwCaseSourcesRPT.CaseID, vwCaseSourcesRPT.Case_Letter, vwCaseSourcesRPT.Number_, vwCaseSourcesRPT.Orig_Atty, vwCaseSourcesRPT.Matter_type, vwCaseSourcesRPT.CaseOpenDate, vwCaseSourcesRPT.CaseNo, vwCaseSourcesRPT.yr, vwCaseSourcesRPT.[Total Earned Fee], vwCaseSourcesRPT.Clsdate, vwCaseSourcesRPT.Closed

FROM vwCaseSourcesRPT;
```
### qryCaseSourcesRPT1
```sql
SELECT qryCaseSourcesRPT.CaseID, TblCase.[Individual Referrer], TblCase.Referral, qryCaseSourcesRPT.[Total Earned Fee] AS Expr2, qryCaseSourcesRPT.Case_Letter, qryCaseSourcesRPT.Number_, qryCaseSourcesRPT.Orig_Atty, qryCaseSourcesRPT.Matter_type, qryCaseSourcesRPT.CaseOpenDate, qryCaseSourcesRPT.CaseNo, qryCaseSourcesRPT.yr, TblCase.Last_Name, TblCase.First_Name, tblDropD.codeval, qryCaseSourcesRPT.Clsdate

FROM (TblCase INNER JOIN qryCaseSourcesRPT ON TblCase.CaseID = qryCaseSourcesRPT.CaseID) ...
```
### qryCaseSourcesRPT_OLD
```sql
SELECT tblCase.CaseID, tblCase.Case_Letter, tblCase.Number_, tblCase.Orig_Atty, tblCase.Matter_type, tblCase.CaseOpenDate, [case_Letter] & [yr] & "-" & [Number_] & "-" & [Orig_Atty] AS CaseNo, tblCase.yr, Disposition.[Total Earned Fee], tblCase.Clsdate, tblCase.Closed

FROM tblCase INNER JOIN Disposition ON tblCase.CaseID = Disposition.CaseID

WHERE (((tblCase.Closed)=Yes));
```
### qryClosing RPT1
```sql
SELECT TblCase.CaseID, TblCase.Last_Name, TblCase.First_Name, TblCase.Address, TblCase.City, TblCase.State, TblCase.Zip, TblCase.HmPhone, TblCase.OtherPhone, TblCase.Closed, TblCase.Clsdate, TblCase.Referral, TblCase.Email, TblCase.DOB, TblCase.SSN, TblCase.Comments, qryClosingRPT.CaseNo, qryClosingRPT.Disposition, qryClosingRPT.[PI Settlement Amount], qryClosingRPT.Dispo_Date, qryClosingRPT.Dispo_Atty, qryClosingRPT.DispoJudge, qryClosingRPT.DispoOppC, qryClosingRPT.Date2, qryClosingRPT.Pay_Out...
```
### qryClosingRPT
```sql
SELECT tblCase.CaseID, [case_Letter] & [yr] & "-" & [Number_] & "-" & [Orig_Atty] AS CaseNo, Disposition.Disposition, Disposition.[PI Settlement Amount], Disposition.Dispo_Date, Disposition.Dispo_Atty, Disposition.DispoJudge, Disposition.DispoOppC, [Matter and AR].Date2, [Matter and AR].Pay_Outlay, [Matter and AR].Charge, [Matter and AR].Payment, [Trust Account].TDate, [Trust Account].TMatter, [Trust Account].Debit, [Trust Account].Credit, tblCase.Matter_type, tblCase.CaseOpenDate, tblCase.Court...
```
### qryCmbCaseClientFile
```sql
SELECT TblCase.CaseID, [TblCase].Last_Name & ", " & [TblCase].First_Name AS Name, Replace([Case_Letter] & [yr] & "-" & [Number_] & "-" & [Orig_Atty],"__","_") AS FileNumber

FROM TblCase INNER JOIN [TB Time Keeping] ON TblCase.CaseID = [TB Time Keeping].CaseID

GROUP BY TblCase.CaseID, Replace([Case_Letter] & [yr] & "-" & [Number_] & "-" & [Orig_Atty],"__","_"), [TblCase].Last_Name & ", " & [TblCase].Last_Name, TblCase.First_Name, TblCase.Last_Name

ORDER BY TblCase.First_Name, TblCase.Last_Name...
```
### qryCmbCaseClientFileFamilyLaw
```sql
SELECT TblCase.CaseID, [TblCase].Last_Name & ", " & [TblCase].First_Name AS Name, Replace([Case_Letter] & [yr] & "-" & [Number_] & "-" & [Orig_Atty],"__","_") AS FileNumber

FROM (TblCase RIGHT JOIN [Family Law - Divorce] ON TblCase.CaseID = [Family Law - Divorce].CaseID) LEFT JOIN [TB Time Keeping] ON TblCase.CaseID = [TB Time Keeping].CaseID

GROUP BY TblCase.CaseID, Replace([Case_Letter] & [yr] & "-" & [Number_] & "-" & [Orig_Atty],"__","_"), [TblCase].Last_Name & ", " & [TblCase].Last_Name, ...
```
### qryCostReimb
```sql
SELECT [Trust Account].CaseID, [Trust Account].TMatter, [Trust Account].Credit

FROM [Trust Account]

WHERE ((([Trust Account].TMatter) Like "*Cost Reimb*"));
```
### qryCostReimbSUM
```sql
SELECT vwCostReimbSUM.CaseID, vwCostReimbSUM.SumOfCredit

FROM vwCostReimbSUM

GROUP BY vwCostReimbSUM.CaseID, vwCostReimbSUM.SumOfCredit;
```
### qryCostReimbSUM_OLD
```sql
SELECT qryCostReimb.CaseID, Sum(qryCostReimb.Credit) AS SumOfCredit

FROM qryCostReimb

GROUP BY qryCostReimb.CaseID;
```
### qryCrimStatus
```sql
SELECT tblCase.Last_Name, tblCase.First_Name, tblCase.Closed, tblCase.Case_Letter, tblCase.yr, tblCase.Number_, tblCase.Orig_Atty, tblCase.Retainer, tblCase.Matter_type, tblCase.Court, tblCase.CType, tblDropD.FieldName, tblCase.CaseID, tblDropD.SortOrder

FROM tblCase LEFT JOIN tblDropD ON tblCase.Orig_Atty = tblDropD.CodeVal

WHERE (((tblCase.Closed)=No) AND ((tblCase.Case_Letter)="C" Or (tblCase.Case_Letter)="T") AND ((tblDropD.FieldName)="orig_atty"));
```
### qryDispoFilter
```sql
SELECT *

FROM Disposition

WHERE (((Disposition.Disposition) Like "n/p" And (Disposition.Disposition) Not Like "$" And (Disposition.Disposition) Not Like "/"));
```
### qryDispos
```sql
SELECT vwDispos.CaseID, vwDispos.Case_Letter, vwDispos.Orig_Atty, vwDispos.Matter_type, vwDispos.Court, vwDispos.CaseOpenDate, vwDispos.HandlingAtty_Case, vwDispos.Dispo_Atty, vwDispos.Dispo_Date, vwDispos.[PI Settlement Amount], vwDispos.[Entire np], vwDispos.[Not Guilty Dismissed], vwDispos.Plea, vwDispos.Trial, vwDispos.Disposition, vwDispos.Name, vwDispos.[Case No], vwDispos.Litigation, vwDispos.CodeVal, vwDispos.FieldName

FROM vwDispos

ORDER BY vwDispos.Dispo_Date DESC;
```
### qryDispos1
```sql
SELECT TblCase.CaseID, TblCase.Case_Letter, TblCase.Orig_Atty, TblCase.Matter_type, TblCase.Court, TblCase.CaseOpenDate, TblCase.HandlingAtty_Case, Disposition.Dispo_Atty, Disposition.Dispo_Date, Disposition.[PI Settlement Amount], Disposition.[Entire np], Disposition.[Not Guilty Dismissed], Disposition.Plea, Disposition.Trial, Disposition.Disposition, [Last_Name] & ", " & [First_Name] AS Name, [case_Letter] & [yr] & "-" & [Number_] & "-" & [Orig_Atty] AS [Case No], [Personal Injury].Litigation,...
```
### qryDispos_OLD
```sql
SELECT TblCase.CaseID, TblCase.Case_Letter, TblCase.Orig_Atty, TblCase.Matter_type, TblCase.Court, TblCase.CaseOpenDate, TblCase.HandlingAtty_Case, Disposition.Dispo_Atty, Disposition.Dispo_Date, Disposition.[PI Settlement Amount], Disposition.[Entire np], Disposition.[Not Guilty Dismissed], Disposition.Plea, Disposition.Trial, Disposition.Disposition, [Last_Name] & ", " & [First_Name] AS Name, [case_Letter] & [yr] & "-" & [Number_] & "-" & [Orig_Atty] AS [Case No], [Personal Injury].Litigation,...
```
### qryEarnedAdvLegal
```sql
SELECT [Trust Account].CaseID, [Trust Account].TMatter, [Trust Account].Credit

FROM [Trust Account]

WHERE ((([Trust Account].TMatter) Like "*adv*"));
```
### qryEarnedAdvLegalSUM
```sql
SELECT vwEarnedAdvLegalSUM.CaseID, vwEarnedAdvLegalSUM.SumOfCredit

FROM vwEarnedAdvLegalSUM;
```
### qryEarnedAdvLegalSUM_OLD
```sql
SELECT qryEarnedAdvLegal.CaseID, Sum(qryEarnedAdvLegal.Credit) AS SumOfCredit

FROM qryEarnedAdvLegal

GROUP BY qryEarnedAdvLegal.CaseID;
```
### qryFamilyLaw
```sql
SELECT [Family Law - Divorce].*, tblCase.*

FROM tblCase RIGHT JOIN [Family Law - Divorce] ON tblCase.CaseID = [Family Law - Divorce].CaseID;
```
### qryFileFolderLabel
```sql
SELECT tblCase.Last_Name, tblCase.First_Name, tblCase.Case_Letter, tblCase.yr, tblCase.Number_, tblCase.Orig_Atty, tblCase.Matter_type, tblCase.HmPhone, tblCase.Court, tblCase.CType, tblCase.CaseID, First(tblHearingDate.Hearing_Date) AS FirstOfHearing_Date, tblCase.Email

FROM tblCase LEFT JOIN tblHearingDate ON tblCase.CaseID = tblHearingDate.CaseID

GROUP BY tblCase.Last_Name, tblCase.First_Name, tblCase.Case_Letter, tblCase.yr, tblCase.Number_, tblCase.Orig_Atty, tblCase.Matter_type, tblCase....
```
### qryInvoiceAttachComp
```sql
SELECT tblCase.CaseID, tblTimeTableDetail.Bill_ID, tblTimeTableDetail.Tdate, tblTimeTableDetail.Description, tblTimeTableDetail.Tatty, tblTimeTableDetail.Rate, tblTimeTableDetail.Time_, [Last_name] & IIf(IsNull([first_name]),Null," " & [first_name]) AS FullName, Nz([time_],0)*Nz([rate],0) AS Amount, [case_Letter] & [yr] & "-" & [Number_] & "-" & [Orig_Atty] AS CaseNo, tblCase.Last_Name, tblCase.First_Name, tblCase.Case_Letter, tblCase.yr, tblCase.Number_, tblCase.Orig_Atty, tblCase.Address, tblC...
```
### qryInvoiceAttachRPT
```sql
SELECT tblCase.CaseID, [TB Time Keeping].Bill_ID, [TB Time Keeping].[Bill Sent], [TB Time Keeping].[Bill Paid], [TB Time Keeping].[Bill Closed], [TB Time Keeping].[BilL Closed Date], [TB Time Keeping].Discount, [case_Letter] & [yr] & "-" & [Number_] & "-" & [Orig_Atty] AS CaseNo, [TB Time Keeping].IANumber, tblTimeTableDetail.Tdate, tblTimeTableDetail.Description, tblTimeTableDetail.Tatty, tblTimeTableDetail.Rate, tblTimeTableDetail.Time_, [TB Time Keeping].TimeNotes

FROM tblCase INNER JOIN ([T...
```
### qryInvoiceAttachRPT1
```sql
SELECT qryInvoiceAttachRPT.CaseID, qryInvoiceAttachRPT.Bill_ID, TblCase.Last_Name, TblCase.First_Name, qryInvoiceAttachRPT.[Bill Sent], qryInvoiceAttachRPT.[Bill Paid], qryInvoiceAttachRPT.[Bill Closed], qryInvoiceAttachRPT.[BilL Closed Date], qryInvoiceAttachRPT.Discount, qryInvoiceAttachRPT.CaseNo, qryInvoiceAttachRPT.IANumber, qryInvoiceAttachRPT.Tdate, qryInvoiceAttachRPT.Description, qryInvoiceAttachRPT.Tatty, qryInvoiceAttachRPT.Rate, qryInvoiceAttachRPT.Time_, [Last_name] & IIf(IsNull([fi...
```
### qryInvoiceComprehensiveTimeDetail
```sql
SELECT qryInvoiceAttachRPT.CaseID, qryInvoiceAttachRPT.Bill_ID, TblCase.Last_Name, TblCase.First_Name, qryInvoiceAttachRPT.IANumber, qryInvoiceAttachRPT.Tdate, qryInvoiceAttachRPT.Description, qryInvoiceAttachRPT.Tatty, qryInvoiceAttachRPT.Rate, qryInvoiceAttachRPT.Time_, Nz([time_],0)*Nz([rate],0) AS Amount

FROM qryInvoiceAttachRPT INNER JOIN TblCase ON qryInvoiceAttachRPT.CaseID = TblCase.CaseID

ORDER BY qryInvoiceAttachRPT.Tdate;
```
### qryInvoiceComprehensiveTimeDetail2
```sql
SELECT qryInvoiceAttachRPT.CaseID, qryInvoiceAttachRPT.Bill_ID, qryInvoiceAttachRPT.IANumber, qryInvoiceAttachRPT.Tdate, qryInvoiceAttachRPT.Description, qryInvoiceAttachRPT.Tatty, qryInvoiceAttachRPT.Rate, qryInvoiceAttachRPT.Time_, Nz([time_],0)*Nz([rate],0) AS Amount

FROM qryInvoiceAttachRPT

ORDER BY qryInvoiceAttachRPT.Tdate;
```
### qryInvoiceComprehensiveTrust
```sql
SELECT vwInvoiceComprehensiveTrust.TrustAccountID, vwInvoiceComprehensiveTrust.TDate, vwInvoiceComprehensiveTrust.TMatter, vwInvoiceComprehensiveTrust.Debit, vwInvoiceComprehensiveTrust.Credit, vwInvoiceComprehensiveTrust.CaseID

FROM vwInvoiceComprehensiveTrust

ORDER BY vwInvoiceComprehensiveTrust.TDate;
```
### qryInvoiceComprehensiveTrustCredit
```sql
SELECT [Trust Account].TDate, [Trust Account].TMatter, [Trust Account].Credit, [Trust Account].OrderNr, [Trust Account].CaseID, [TB Time Keeping].[BilL Closed Date]

FROM [TB Time Keeping] LEFT JOIN [Trust Account] ON [TB Time Keeping].CaseID = [Trust Account].CaseID

WHERE ((([Trust Account].TMatter) Not Like "*Earned*" And ([Trust Account].TMatter) Not Like "*Reimb*" And ([Trust Account].TMatter) Not Like "*Refund*") AND (([Trust Account].Credit)>0) AND (([Trust Account].CaseID)=9966))

ORDER ...
```
### qryInvoiceComprehensiveTrustCredit2
```sql
SELECT qryInvoiceComprehensiveTrustCredit.TDate, qryInvoiceComprehensiveTrustCredit.TMatter, qryInvoiceComprehensiveTrustCredit.Credit, qryInvoiceComprehensiveTrustCredit.OrderNr, [TB Time Keeping].Bill_ID, [TB Time Keeping].CaseID, [TB Time Keeping].[BilL Closed Date]

FROM [TB Time Keeping] INNER JOIN qryInvoiceComprehensiveTrustCredit ON [TB Time Keeping].CaseID = qryInvoiceComprehensiveTrustCredit.CaseID;
```
### qryInvoiceComprehensiveTrustCredit3
```sql
SELECT qryInvoiceComprehensiveTrustCredit2.TDate, qryInvoiceComprehensiveTrustCredit2.TMatter, qryInvoiceComprehensiveTrustCredit2.Credit, qryInvoiceComprehensiveTrustCredit2.OrderNr, qryInvoiceComprehensiveTrustCredit2.CaseID, [TB Time Keeping].[BilL Closed Date], qryInvoiceComprehensiveTrustCredit2.CaseID, [TB Time Keeping].Bill_ID

FROM qryInvoiceComprehensiveTrustCredit2 RIGHT JOIN [TB Time Keeping] ON qryInvoiceComprehensiveTrustCredit2.Bill_ID = [TB Time Keeping].Bill_ID

WHERE (((qryInvoi...
```
### qryInvoiceComprehensiveTrustCredit4
```sql
SELECT tblCase.CaseID, tblCase.CaseOpenDate, [TB Time Keeping].[BilL Closed Date], [TB Time Keeping].Bill_ID, [Trust Account].TrustAccountID, [Trust Account].TDate, [Trust Account].TMatter, [Trust Account].Debit, [Trust Account].Credit, [Trust Account].OrderNr

FROM (tblCase INNER JOIN [TB Time Keeping] ON tblCase.CaseID = [TB Time Keeping].CaseID) INNER JOIN [Trust Account] ON tblCase.CaseID = [Trust Account].CaseID

WHERE ((([Trust Account].TDate)<[Bill Closed Date] And ([Trust Account].TDate)...
```
### qryInvoiceComprehensiveTrust_OLD
```sql
SELECT qryStmtTrustRPT.TrustAccountID, qryStmtTrustRPT.TDate, qryStmtTrustRPT.TMatter, qryStmtTrustRPT.Debit, qryStmtTrustRPT.Credit, tblCase.CaseID

FROM (qryStmtTrustRPT LEFT JOIN tblCase ON qryStmtTrustRPT.CaseID = tblCase.CaseID) LEFT JOIN qryTrustAccount ON qryStmtTrustRPT.TrustAccountID = qryTrustAccount.TrustAccountID

ORDER BY qryStmtTrustRPT.TDate;
```
### qryInvoiceRPT
```sql
SELECT vwInvoiceRPT.CaseID, vwInvoiceRPT.MatterID, vwInvoiceRPT.Date2, vwInvoiceRPT.Pay_Outlay, vwInvoiceRPT.Charge, vwInvoiceRPT.Payment, vwInvoiceRPT.[Case No], vwInvoiceRPT.ID, vwInvoiceRPT.[Balance Due Date], vwInvoiceRPT.[Past Due], vwInvoiceRPT.[Long Term Collections], vwInvoiceRPT.chkBalanceDue, vwInvoiceRPT.[Billing Notes]

FROM vwInvoiceRPT;
```
### qryInvoiceRPT1
```sql
SELECT vwInvoiceRPT1.CaseID, vwInvoiceRPT1.Last_Name, vwInvoiceRPT1.First_Name, vwInvoiceRPT1.CaseOpenDate, vwInvoiceRPT1.Closed, vwInvoiceRPT1.Clsdate, vwInvoiceRPT1.Extended_Ledger, vwInvoiceRPT1.Case_Letter, vwInvoiceRPT1.yr, vwInvoiceRPT1.Number_, vwInvoiceRPT1.Orig_Atty, vwInvoiceRPT1.Address, vwInvoiceRPT1.CourtCaseNo, vwInvoiceRPT1.City, vwInvoiceRPT1.FamilyLaw, vwInvoiceRPT1.State, vwInvoiceRPT1.Zip, vwInvoiceRPT1.Country, vwInvoiceRPT1.HmPhone, vwInvoiceRPT1.Action, vwInvoiceRPT1.OtherP...
```
### qryInvoiceRPT1_OLD
```sql
SELECT tblCase.*, qryMatter.Balance, qryInvoiceRPT.MatterID, qryInvoiceRPT.Date2, qryInvoiceRPT.Pay_Outlay, qryInvoiceRPT.Charge, qryInvoiceRPT.Payment, qryInvoiceRPT.[Case No], qryInvoiceRPT.[Balance Due Date], qryInvoiceRPT.[Past Due], qryInvoiceRPT.[Long Term Collections], qryInvoiceRPT.chkBalanceDue, qryInvoiceRPT.[Billing Notes], fncRunningDebit([tblcase].[CaseID],qryInvoiceRPT.[Date2],qryInvoiceRPT.[MatterID]) AS RunningDebit, fncRunningCredit([tblcase].[CaseID],qryInvoiceRPT.[Date2],qryIn...
```
### qryInvoiceRPT_OLD
```sql
SELECT [Matter and AR].CaseID, [Matter and AR].MatterID, [Matter and AR].Date2, [Matter and AR].Pay_Outlay, [Matter and AR].Charge, [Matter and AR].Payment, [case_Letter] & [yr] & "-" & [Number_] & "-" & [Orig_Atty] AS [Case No], Billing.ID, Billing.[Balance Due Date], Billing.[Past Due], Billing.[Long Term Collections], Billing.chkBalanceDue, Billing.[Billing Notes]

FROM (tblCase LEFT JOIN Billing ON tblCase.CaseID = Billing.CaseID) LEFT JOIN [Matter and AR] ON tblCase.CaseID = [Matter and AR]...
```
### qryInvoiceTrustCostBillDate
```sql
SELECT qryInvoiceComprehensiveTrustCredit.CaseID, qryInvoiceComprehensiveTrustCredit.TDate, qryInvoiceComprehensiveTrustCredit.TMatter, qryInvoiceComprehensiveTrustCredit.Credit, qryInvoiceComprehensiveTrustCredit.OrderNr, [TB Time Keeping].[BilL Closed Date], qryInvoiceComprehensiveTrustCredit.Bill_ID

FROM [TB Time Keeping] LEFT JOIN qryInvoiceComprehensiveTrustCredit ON [TB Time Keeping].Bill_ID = qryInvoiceComprehensiveTrustCredit.Bill_ID

WHERE (((qryInvoiceComprehensiveTrustCredit.TDate)<[...
```
### qryMatter
```sql
SELECT vwMatter.MatterID, vwMatter.Date2, vwMatter.CaseID, vwMatter.SumOfCharge, vwMatter.SumOfPayment, vwMatter.Balance, vwMatter.OrderNr

FROM vwMatter

ORDER BY vwMatter.CaseID, vwMatter.OrderNr;
```
### qryMatterBalanceTotals
```sql
SELECT vwMatterBalanceTotals.CaseID, vwMatterBalanceTotals.SumOfBalance

FROM vwMatterBalanceTotals;
```
### qryMatterBalanceTotals_OLD
```sql
SELECT qryMatter.CaseID, Sum(qryMatter.Balance) AS SumOfBalance

FROM qryMatter

GROUP BY qryMatter.CaseID;
```
### qryMatterSums
```sql
SELECT qryMatter.CaseID, Sum(qryMatter.SumOfCharge) AS SumOfSumOfCharge, Sum(qryMatter.SumOfPayment) AS SumOfSumOfPayment

FROM qryMatter

GROUP BY qryMatter.CaseID;
```
### qryMatter_OLD
```sql
SELECT [Matter and AR].MatterID, [Matter and AR].Date2, [Matter and AR].CaseID, Sum([Matter and AR].Charge) AS SumOfCharge, Sum([Matter and AR].Payment) AS SumOfPayment, Sum(Nz([Charge],0)-Nz([payment],0)) AS Balance, [Matter and AR].OrderNr

FROM [Matter and AR]

GROUP BY [Matter and AR].MatterID, [Matter and AR].Date2, [Matter and AR].CaseID, [Matter and AR].OrderNr

ORDER BY [Matter and AR].CaseID, [Matter and AR].OrderNr;
```
### qryMergeTest
```sql
SELECT tblCase.Last_Name, tblCase.First_Name, tblCase.Court, tbl_CtCaseNumbers.CtNumber, tblCase.CaseID, tblCase.Closed

FROM tblCase INNER JOIN tbl_CtCaseNumbers ON tblCase.CaseID = tbl_CtCaseNumbers.CaseID

WHERE (((tblCase.Closed)=No));
```
### qryNewInvoice02Comp
```sql
SELECT qryNewInvoice_01.*

FROM qryNewInvoice_01

ORDER BY qryNewInvoice_01.MatterID;
```
### qryNewInvoice_01
```sql
SELECT qryInvoiceRPT1.*, [Runningbalance]+[Retainer] AS RetBal

FROM qryInvoiceRPT1

ORDER BY qryInvoiceRPT1.MatterID DESC;
```
### qryNewInvoice_02
```sql
SELECT qryNewInvoice_01.*

FROM qryNewInvoice_01

ORDER BY qryNewInvoice_01.MatterID;
```
### qryNewTrustComp
```sql
SELECT vwNewTrustComp.CaseID, vwNewTrustComp.Last_Name, vwNewTrustComp.First_Name, vwNewTrustComp.CaseOpenDate, vwNewTrustComp.Closed, vwNewTrustComp.Clsdate, vwNewTrustComp.Extended_Ledger, vwNewTrustComp.Case_Letter, vwNewTrustComp.yr, vwNewTrustComp.Number_, vwNewTrustComp.Orig_Atty, vwNewTrustComp.Address, vwNewTrustComp.CourtCaseNo, vwNewTrustComp.City, vwNewTrustComp.FamilyLaw, vwNewTrustComp.State, vwNewTrustComp.Zip, vwNewTrustComp.Country, vwNewTrustComp.HmPhone, vwNewTrustComp.Action, ...
```
### qryNewTrustComp_OLD
```sql
SELECT tblCase.*, qryTrustAccount.SumOfDebit, qryTrustAccount.SumOfCredit, qryTrustAccount.Balance, qryStmtTrustRPT.TrustAccountID, qryStmtTrustRPT.TDate, qryStmtTrustRPT.TMatter, qryStmtTrustRPT.Debit, qryStmtTrustRPT.Credit, qryStmtTrustRPT.CheckNumber, qryStmtTrustRPT.[Case No], qryStmtTrustRPT.CheckCashed, qryStmtTrustRPT.DepCleared, qryStmtTrustRPT.Reconciled, qryTrustAccount.OrderNr, [tblcase.Last_Name] & ", " & [tblcase.First_Name] AS Name, Replace([tblcase.Case_Letter] & [tblcase.yr] & "...
```
### qryOutstandingARRPT
```sql
SELECT tblCase.CaseID, tblCase.Case_Letter, tblCase.yr, tblCase.Number_, tblCase.Orig_Atty, tblCase.Matter_type, tblCase.Retainer, tblCase.CaseOpenDate, qryMatter.Balance, Billing.[Balance Due Date], Billing.[Past Due], Billing.[Long Term Collections], Billing.chkBalanceDue, [case_Letter] & [yr] & "-" & [Number_] & "-" & [Orig_Atty] AS CaseNo, Billing.WriteOff

FROM (tblCase INNER JOIN qryMatter ON tblCase.CaseID = qryMatter.CaseID) INNER JOIN Billing ON tblCase.CaseID = Billing.CaseID;
```
### qryOutstandingARRPT1
```sql
SELECT TblCase.CaseID, TblCase.Last_Name, TblCase.First_Name, qryOutstandingARRPT.CaseNo, qryOutstandingARRPT.Matter_type, qryOutstandingARRPT.Retainer, qryOutstandingARRPT.CaseOpenDate, qryOutstandingARRPT.Balance, qryOutstandingARRPT.[Balance Due Date], qryOutstandingARRPT.[Past Due], qryOutstandingARRPT.[Long Term Collections], qryOutstandingARRPT.chkBalanceDue, [qryOutstandingARRPT.Retainer]+[Balance] AS RetBal, qryOutstandingARRPT.Case_Letter, qryOutstandingARRPT.Orig_Atty, qryOutstandingAR...
```
### qryPersInjStatus
```sql
SELECT tblCase.Last_Name, tblCase.First_Name, tblCase.Closed, tblCase.Case_Letter, tblCase.yr, tblCase.Number_, tblCase.Orig_Atty, tblCase.CourtCaseNo, tblCase.SOL, [Personal Injury].CaseID, [Personal Injury].[Filing Date], [Personal Injury].Litigation, [Personal Injury].DOI, [Personal Injury].Demand, [Personal Injury].BriefDescription, [Personal Injury].ID, tblDropD.SortOrder, tblDropD.FieldName, [Personal Injury].ServedDate, [Personal Injury].CompltServed, [Personal Injury].PISOL, [Personal In...
```
### qryReceipt
```sql
SELECT tblCase.Last_Name, tblCase.First_Name, tblCase.Matter_type, tblCase.Case_Letter, tblCase.yr, tblCase.Number_, tblCase.Orig_Atty, [Matter and AR].CaseID, [Matter and AR].Date2, [Matter and AR].Pay_Outlay, [Matter and AR].Payment, [Matter and AR].MatterID, tblCase.Retainer

FROM tblCase INNER JOIN [Matter and AR] ON tblCase.CaseID = [Matter and AR].CaseID

WHERE ((([Matter and AR].Payment)>0));
```
### qryReconciliationWFBankBalance
```sql
SELECT;
```
### qryReconciliation_sumOfBalances
```sql
SELECT Sum(qryTakeOff_trust_account.Balance) AS SumOfBalance

FROM qryTakeOff_trust_account;
```
### qryReconciliation_sumOfCredit
```sql
SELECT Sum(qryTakeOff_unchashed_checks.SumOfCredit) AS SumOfCredit

FROM qryTakeOff_unchashed_checks;
```
### qryReconciliation_sumOfUnclearedDeposits
```sql
SELECT Sum(qryTakeOff_uncleared_deposits.SumOfDebit) AS SumOfDebit

FROM qryTakeOff_uncleared_deposits;
```
### qryRunningSum
```sql
SELECT qryInvoiceRPT1.CaseID, qryInvoiceRPT1.MatterID, qryInvoiceRPT1.Charge, qryInvoiceRPT1.Payment, qryInvoiceRPT1.CaseOpenDate, qryInvoiceRPT1.Date2, fncRunningDebit([CaseID],[Date2],[MatterID]) AS RunningDebit, fncRunningCredit([CaseID],[Date2],[MatterID]) AS RunningCredit, [RunningDebit]-[RunningCredit] AS RunningBalance

FROM qryInvoiceRPT1

WHERE (((qryInvoiceRPT1.CaseID)=11));
```
### qrySOL
```sql
SELECT [Personal Injury].CaseID, tblCase.Last_Name, tblCase.First_Name, tblCase.Action_Needed_on_Payment, tblCase.Action, tblCase.SOL, tblCase.Comments, tblCase.Case_Letter, tblCase.yr, tblCase.Number_, tblCase.Orig_Atty, tblCase.Matter_type, tblCase.CaseOpenDate, tblCase.HandlingAtty_Case, [Personal Injury].[Filing Date], [Last_Name] & ", " & [First_Name] AS Name, [case_Letter] & [yr] & "-" & [Number_] & "-" & [Orig_Atty] AS [Case No], tblCase.ParaLegal, [Personal Injury].DOI, tblCase.Closed

F...
```
### qryStmtTrustRPT
```sql
SELECT vwStmtTrustRPT.CaseID, vwStmtTrustRPT.Case_Letter, vwStmtTrustRPT.yr, vwStmtTrustRPT.Number_, vwStmtTrustRPT.Orig_Atty, vwStmtTrustRPT.Matter_type, vwStmtTrustRPT.Retainer, vwStmtTrustRPT.CaseOpenDate, vwStmtTrustRPT.TrustAccountID, vwStmtTrustRPT.TDate, vwStmtTrustRPT.TMatter, vwStmtTrustRPT.Debit, vwStmtTrustRPT.Credit, vwStmtTrustRPT.CheckCashed, vwStmtTrustRPT.CheckNumber, vwStmtTrustRPT.[Case No], vwStmtTrustRPT.DepCleared, vwStmtTrustRPT.Reconciled, vwStmtTrustRPT.OrderNr

FROM vwSt...
```
### qryStmtTrustRPT1
```sql
SELECT vwStmtTrustRPT1.*

FROM vwStmtTrustRPT1

ORDER BY vwStmtTrustRPT1.TrustAccountID, vwStmtTrustRPT1.OrderNr;
```
### qryStmtTrustRPT1_OLD
```sql
SELECT tblCase.*, qryTrustAccount.SumOfDebit, qryTrustAccount.SumOfCredit, qryTrustAccount.Balance, qryStmtTrustRPT.TrustAccountID, qryStmtTrustRPT.TDate, qryStmtTrustRPT.TMatter, qryStmtTrustRPT.Debit, qryStmtTrustRPT.Credit, qryStmtTrustRPT.CheckNumber, qryStmtTrustRPT.[Case No], qryStmtTrustRPT.CheckCashed, qryStmtTrustRPT.DepCleared, qryStmtTrustRPT.Reconciled, qryTrustAccount.OrderNr, [tblcase.Last_Name] & ", " & [tblcase.First_Name] AS Name, Replace([tblcase.Case_Letter] & [tblcase.yr] & "...
```
### qryStmtTrustRPT_OLD
```sql
SELECT tblCase.CaseID, tblCase.Case_Letter, tblCase.yr, tblCase.Number_, tblCase.Orig_Atty, tblCase.Matter_type, tblCase.Retainer, tblCase.CaseOpenDate, [Trust Account].TrustAccountID, [Trust Account].TDate, [Trust Account].TMatter, [Trust Account].Debit, [Trust Account].Credit, [Trust Account].CheckCashed, [Trust Account].CheckNumber, [case_Letter] & [yr] & "-" & [Number_] & "-" & [Orig_Atty] AS [Case No], [Trust Account].DepCleared, [Trust Account].Reconciled, [Trust Account].OrderNr

FROM tbl...
```
### qrySumofPayments
```sql
SELECT [Matter and AR].CaseID, Sum([Matter and AR].Payment) AS SumOfPayment

FROM [Matter and AR]

WHERE ((([Matter and AR].Pay_Outlay) Like "*payment*"))

GROUP BY [Matter and AR].CaseID;
```
### qryTKClose
```sql
SELECT vwTKClose_A.CaseID, vwTKClose_A.FileNumber, vwTKClose_A.Name, vwTKClose_A.Orig_Atty, vwTKClose_A.HandlingAtty_Case, vwTKClose_A.SumOfAdvancedAR, vwTKClose_A.CostHold, vwTKClose_A.SumOfUnclearedDeposits, vwTKClose_A.Balance, vwTKClose_A.SumOfUncashedChecks, vwTKClose_A.AvailBalance, vwTKClose_A.BankBalance, vwTKClose_A.SumOfTotal, vwTKClose_A.IANumber, vwTKClose_A.Bill_ID, vwTKClose_A.Retainer, vwTKClose_A.RetainerReimb, vwTKClose_A.RetReimbAmount, vwTKClose_A.MaxOfMatterID

FROM vwTKClose...
```
### qryTKClose1
```sql
SELECT qryTKClose_A.CaseID, qryTKClose_A.FileNumber, qryTKClose_A.Name, qryTKClose_A.Orig_Atty, qryTKClose_A.HandlingAtty_Case, qryTKClose_A.SumOfAdvancedAR, qryTKClose_A.CostHold, qryTKClose_A.SumOfUnclearedDeposits, qryTKClose_A.Balance, qryTKClose_A.SumOfUncashedChecks, [Balance]-Nz([SumOfUnclearedDeposits],0) AS AvailBalance, qryTKClose_A.SumOfTotal, [Balance]+Nz([SumOfUncashedChecks],0)-Nz([SumOfUnclearedDeposits],0) AS BankBalance, qryTKClose_A.IANumber, qryTKClose_A.Bill_ID, qryTKClose_A....
```
### qryTKClose_A
```sql
SELECT vwTKClose_A.CaseID, vwTKClose_A.FileNumber, vwTKClose_A.Name, vwTKClose_A.Orig_Atty, vwTKClose_A.HandlingAtty_Case, vwTKClose_A.SumOfAdvancedAR, vwTKClose_A.CostHold, vwTKClose_A.SumOfUnclearedDeposits, vwTKClose_A.Balance, vwTKClose_A.SumOfUncashedChecks, vwTKClose_A.AvailBalance, vwTKClose_A.BankBalance, vwTKClose_A.SumOfTotal, vwTKClose_A.IANumber, vwTKClose_A.Bill_ID, vwTKClose_A.Retainer, vwTKClose_A.RetainerReimb, vwTKClose_A.RetReimbAmount, vwTKClose_A.MaxOfMatterID, vwTKClose_A.AR...
```
### qryTKClose_A_OLD
```sql
SELECT tblCase.CaseID, Replace([Case_Letter] & [yr] & "-" & [Number_] & "-" & [Orig_Atty],"__","_") AS FileNumber, [Last_Name] & ", " & [First_Name] AS Name, tblCase.Orig_Atty, tblCase.HandlingAtty_Case, qryTakeOff_advanced_AR.SumOfCharge AS SumOfAdvancedAR, tblCase.CostHold, qryTakeOff_uncleared_deposits.SumOfDebit AS SumOfUnclearedDeposits, qryTakeOff_trust_account.Balance, qryTakeOff_unchashed_checks.SumOfCredit AS SumOfUncashedChecks, [Balance]-Nz([SumOfUnclearedDeposits],0) AS AvailBalance,...
```
### qryTKClose_OLD
```sql
SELECT tblCase.CaseID, Replace([Case_Letter] & [yr] & "-" & [Number_] & "-" & [Orig_Atty],"__","_") AS FileNumber, [Last_Name] & ", " & [First_Name] AS Name, tblCase.Orig_Atty, tblCase.HandlingAtty_Case, qryTakeOff_advanced_AR.SumOfCharge AS SumOfAdvancedAR, qryTakeOff_cost_hold.CostHold, qryTakeOff_uncleared_deposits.SumOfDebit AS SumOfUnclearedDeposits, qryTakeOff_trust_account.Balance, qryTakeOff_unchashed_checks.SumOfCredit AS SumOfUncashedChecks, [Balance]-Nz([SumOfUnclearedDeposits],0) AS ...
```
### qryTTAmount
```sql
SELECT tblTimeTableDetail.Time_ID, tblTimeTableDetail.Bill_ID, tblTimeTableDetail.Time_, tblTimeTableDetail.Rate, Nz([Time_],0)*Nz([rate],0) AS Amount

FROM tblTimeTableDetail;
```
### qryTTAmountAtty
```sql
SELECT tblTimeTableDetail.Time_ID, tblTimeTableDetail.Bill_ID, tblTimeTableDetail.Tatty, tblTimeTableDetail.Rate, tblTimeTableDetail.Time_, Nz([Time_],0)*Nz([rate],0) AS Amount

FROM tblTimeTableDetail;
```
### qryTTAmountHours
```sql
SELECT tblTimeTableDetail.Time_ID, tblTimeTableDetail.Bill_ID, tblTimeTableDetail.Tatty, tblTimeTableDetail.Rate, tblTimeTableDetail.Time_, Sum(Nz([Time_],0)) AS Amount

FROM tblTimeTableDetail;
```
### qryTTAmountHours_SUM
```sql
SELECT tblTimeTableDetail.Bill_ID, Sum(tblTimeTableDetail.Time_) AS SumOfTime_

FROM tblTimeTableDetail

GROUP BY tblTimeTableDetail.Bill_ID;
```
### qryTTAmountHours_SUM_byAtty
```sql
SELECT tblTimeTableDetail.Bill_ID, Sum(tblTimeTableDetail.Time_) AS SumOfTime_, tblTimeTableDetail.Tatty

FROM tblTimeTableDetail

GROUP BY tblTimeTableDetail.Bill_ID, tblTimeTableDetail.Tatty;
```
### qryTTAmountHours_SUM_byAtty_TotalCaseID
```sql
SELECT vwTTAmountHours_SUM_byAtty_TotalCaseID.SumOfSumOfTime_, vwTTAmountHours_SUM_byAtty_TotalCaseID.Tatty, vwTTAmountHours_SUM_byAtty_TotalCaseID.CaseID

FROM vwTTAmountHours_SUM_byAtty_TotalCaseID

GROUP BY vwTTAmountHours_SUM_byAtty_TotalCaseID.SumOfSumOfTime_, vwTTAmountHours_SUM_byAtty_TotalCaseID.Tatty, vwTTAmountHours_SUM_byAtty_TotalCaseID.CaseID;
```
### qryTTAmountHours_SUM_byAtty_TotalCaseID_OLD
```sql
SELECT Sum(qryTTAmountHours_SUM_byAtty.SumOfTime_) AS SumOfSumOfTime_, qryTTAmountHours_SUM_byAtty.Tatty, [TB Time Keeping].CaseID

FROM qryTTAmountHours_SUM_byAtty LEFT JOIN [TB Time Keeping] ON qryTTAmountHours_SUM_byAtty.Bill_ID = [TB Time Keeping].Bill_ID

GROUP BY qryTTAmountHours_SUM_byAtty.Tatty, [TB Time Keeping].CaseID;
```
### qryTTAmountHours_TotalCaseID
```sql
SELECT vwTTAmountHours_TotalCaseID.CaseID, vwTTAmountHours_TotalCaseID.SumOfSumOfTime_

FROM vwTTAmountHours_TotalCaseID

GROUP BY vwTTAmountHours_TotalCaseID.CaseID, vwTTAmountHours_TotalCaseID.SumOfSumOfTime_;
```
### qryTTAmountHours_TotalCaseID_OLD
```sql
SELECT Sum(qryTTAmountHours_SUM.SumOfTime_) AS SumOfSumOfTime_, [TB Time Keeping].CaseID

FROM qryTTAmountHours_SUM INNER JOIN [TB Time Keeping] ON qryTTAmountHours_SUM.Bill_ID = [TB Time Keeping].Bill_ID

GROUP BY [TB Time Keeping].CaseID;
```
### qryTakeOff
```sql
SELECT qryTakeOff_A.CaseID, qryTakeOff_A.FileNumber, qryTakeOff_A.Name, qryTakeOff_A.Last_Name, qryTakeOff_A.Orig_Atty, qryTakeOff_A.Matter_type, qryTakeOff_A.HandlingAtty_Case, qryTakeOff_A.SumOfAdvancedAR, qryTakeOff_A.CostHold, qryTakeOff_A.SumOfUnclearedDeposits, qryTakeOff_A.Balance, qryTakeOff_A.SumOfUncashedChecks, [Balance]-Nz([SumOfUnclearedDeposits],0) AS AvailBalance, [Balance]+Nz([SumOfUncashedChecks],0)-Nz([SumOfUnclearedDeposits],0) AS BankBalance, qryTakeOff_A.SumOfTotal, qryTakeO...
```
### qryTakeOff2
```sql
SELECT tblCase.CaseID, Replace([Case_Letter] & [yr] & "-" & [Number_] & "-" & [Orig_Atty],"__","_") AS FileNumber, [Last_Name] & ", " & [First_Name] AS Name, tblCase.Orig_Atty, tblCase.HandlingAtty_Case, qryTakeOff_advanced_AR.SumOfCharge AS SumOfAdvancedAR, qryTakeOff_cost_hold.CostHold, qryTakeOff_uncleared_deposits.SumOfDebit AS SumOfUnclearedDeposits, qryTakeOff_trust_account.Balance, qryTakeOff_unchashed_checks.SumOfCredit AS SumOfUncashedChecks, [Balance]-Nz([SumOfUnclearedDeposits],0) AS ...
```
### qryTakeOffDate
```sql
SELECT tblTakeOffMonth.TakeOffDate

FROM tblTakeOffMonth INNER JOIN tblTakeOff ON tblTakeOffMonth.TakeOffMonthID = tblTakeOff.TakeOffMonthID;
```
### qryTakeOffStep2
```sql
SELECT vwTakeOffStep2.FileNumber, vwTakeOffStep2.Name, vwTakeOffStep2.CaseID, vwTakeOffStep2.Last_Name, vwTakeOffStep2.First_Name, vwTakeOffStep2.CaseOpenDate, vwTakeOffStep2.Closed, vwTakeOffStep2.Clsdate, vwTakeOffStep2.Extended_Ledger, vwTakeOffStep2.Case_Letter, vwTakeOffStep2.yr, vwTakeOffStep2.Number_, vwTakeOffStep2.Orig_Atty, vwTakeOffStep2.Address, vwTakeOffStep2.CourtCaseNo, vwTakeOffStep2.City, vwTakeOffStep2.FamilyLaw, vwTakeOffStep2.State, vwTakeOffStep2.Zip, vwTakeOffStep2.Country,...
```
### qryTakeOffStep2_OLD
```sql
SELECT Replace([Case_Letter] & [yr] & "-" & [Number_] & "-" & [Orig_Atty],"__","_") AS FileNumber, [Last_Name] & ", " & [First_Name] AS Name, tblCase.*, tblTakeOff.*

FROM tblCase LEFT JOIN tblTakeOff ON tblCase.CaseID = tblTakeOff.CaseID

ORDER BY [Last_Name] & ", " & [First_Name];
```
### qryTakeOff_A
```sql
SELECT vwTakeOff_A.CaseID, vwTakeOff_A.FileNumber, vwTakeOff_A.Name, vwTakeOff_A.Last_Name, vwTakeOff_A.Orig_Atty, vwTakeOff_A.HandlingAtty_Case, vwTakeOff_A.SumOfAdvancedAR, vwTakeOff_A.SumOfUnclearedDeposits, vwTakeOff_A.Balance, vwTakeOff_A.SumOfUncashedChecks, vwTakeOff_A.SumOfTotal, vwTakeOff_A.IANumber, vwTakeOff_A.Bill_ID, vwTakeOff_A.CostHold, vwTakeOff_A.SumOfCostBalance, vwTakeOff_A.SumofPrepaid, vwTakeOff_A.SumAdvLegal, vwTakeOff_A.SumEarnedAdv, vwTakeOff_A.SumCostReimb, vwTakeOff_A.M...
```
### qryTakeOff_A_OLD
```sql
SELECT tblCase.CaseID, Replace([Case_Letter] & [yr] & "-" & [Number_] & "-" & [Orig_Atty],"__","_") AS FileNumber, [Last_Name] & ", " & [First_Name] AS Name, tblCase.Orig_Atty, tblCase.HandlingAtty_Case, qryTakeOff_advanced_AR.SumOfCharge AS SumOfAdvancedAR, qryTakeOff_uncleared_deposits.SumOfDebit AS SumOfUnclearedDeposits, qryTakeOff_trust_account.Balance, qryTakeOff_unchashed_checks.SumOfCredit AS SumOfUncashedChecks, qry_TimeKeeping_CaseID_totals.SumOfTotal, qry_TimeKeeping_CaseID_totals.IAN...
```
### qryTakeOff_advanced_AR
```sql
SELECT tblCase.CaseID, Sum([Matter and AR].Charge) AS SumOfCharge

FROM tblCase INNER JOIN [Matter and AR] ON tblCase.CaseID = [Matter and AR].CaseID

GROUP BY tblCase.CaseID, [Matter and AR].FirmPrepaid

HAVING ((([Matter and AR].FirmPrepaid)=Yes));
```
### qryTakeOff_cost_hold
```sql
SELECT tblCase.CaseID, tblCase.CostHold

FROM tblCase;
```
### qryTakeOff_trust_account
```sql
SELECT vwTakeOff_trust_Account.CaseID, vwTakeOff_trust_Account.SumOfDebit, vwTakeOff_trust_Account.SumOfCredit, vwTakeOff_trust_Account.Balance

FROM vwTakeOff_trust_Account

ORDER BY vwTakeOff_trust_Account.CaseID;
```
### qryTakeOff_trust_account_OLD
```sql
SELECT [Trust Account].CaseID, Sum([Trust Account].Debit) AS SumOfDebit, Sum([Trust Account].Credit) AS SumOfCredit, Sum(Nz([debit],0)-Nz([Credit],0)) AS Balance

FROM [Trust Account]

GROUP BY [Trust Account].CaseID;
```
### qryTakeOff_unchashed_checks
```sql
SELECT tblCase.CaseID, Sum([Trust Account].Credit) AS SumOfCredit

FROM tblCase INNER JOIN [Trust Account] ON tblCase.CaseID = [Trust Account].CaseID

GROUP BY tblCase.CaseID, [Trust Account].CheckCashed

HAVING ((([Trust Account].CheckCashed)=Yes));
```
### qryTakeOff_uncleared_deposits
```sql
SELECT tblCase.CaseID, Sum([Trust Account].Debit) AS SumOfDebit

FROM tblCase INNER JOIN [Trust Account] ON tblCase.CaseID = [Trust Account].CaseID

GROUP BY tblCase.CaseID, [Trust Account].DepCleared

HAVING ((([Trust Account].DepCleared)=Yes));
```
### qryTimeKeeping
```sql
SELECT [TB Time Keeping].Bill_ID, [TB Time Keeping].[Bill Sent], [TB Time Keeping].[Bill Paid], [TB Time Keeping].[Bill Closed], [TB Time Keeping].[BilL Closed Date], [TB Time Keeping].Discount, [TB Time Keeping].[Bill Number], [TB Time Keeping].IANumber, [TB Time Keeping].[Bill Open], [TB Time Keeping].TimeNotes, tblCase.*, [TB Time Keeping].TKLocked, [TB Time Keeping].InvoiceTotalAdvance, [TB Time Keeping].InvoiceExceedsTrust, [TB Time Keeping].StatementLessTrust, [TB Time Keeping].TrustatClos...
```
### qryTimeKeepingClosed
```sql
SELECT vwTimeKeepingClosed.[Bill Sent], vwTimeKeepingClosed.[Bill Paid], vwTimeKeepingClosed.[Bill Closed], vwTimeKeepingClosed.[BilL Closed Date], vwTimeKeepingClosed.Discount, vwTimeKeepingClosed.IANumber, vwTimeKeepingClosed.FileNumber, vwTimeKeepingClosed.BalanceCalculated, vwTimeKeepingClosed.CaseID, vwTimeKeepingClosed.Last_Name, vwTimeKeepingClosed.First_Name, vwTimeKeepingClosed.CaseOpenDate, vwTimeKeepingClosed.Closed, vwTimeKeepingClosed.Clsdate, vwTimeKeepingClosed.Extended_Ledger, vw...
```
### qryTimeKeepingClosed_Old
```sql
SELECT [TB Time Keeping].[Bill Sent], [TB Time Keeping].[Bill Paid], [TB Time Keeping].[Bill Closed], [TB Time Keeping].[BilL Closed Date], [TB Time Keeping].Discount, [TB Time Keeping].IANumber, Replace([Case_Letter] & [yr] & "-" & [Number_] & "-" & [Orig_Atty],"__","_") AS FileNumber, [SumOfAmount]-Nz([Discount]) AS BalanceCalculated, tblCase.*, [TB Time Keeping].[Bill Open], [tblCase].[Last_Name] & ", " & [tblCase].[First_Name] AS Name, [TB Time Keeping].Bill_ID, [TB Time Keeping].TrustatClos...
```
### qryTimeKeepingOpen
```sql
SELECT vwTimeKeepingOpen.[Bill Closed], vwTimeKeepingOpen.Bill_ID, vwTimeKeepingOpen.IANumber, vwTimeKeepingOpen.FileNumber, vwTimeKeepingOpen.BalanceCalculated, vwTimeKeepingOpen.CaseID, vwTimeKeepingOpen.Last_Name, vwTimeKeepingOpen.First_Name, vwTimeKeepingOpen.CaseOpenDate, vwTimeKeepingOpen.Closed, vwTimeKeepingOpen.Clsdate, vwTimeKeepingOpen.Extended_Ledger, vwTimeKeepingOpen.Case_Letter, vwTimeKeepingOpen.yr, vwTimeKeepingOpen.Number_, vwTimeKeepingOpen.Orig_Atty, vwTimeKeepingOpen.Addres...
```
### qryTimeKeepingOpen_OLD
```sql
SELECT [TB Time Keeping].[Bill Closed], [TB Time Keeping].Bill_ID, [TB Time Keeping].IANumber, Replace([Case_Letter] & [yr] & "-" & [Number_] & "-" & [Orig_Atty],"__","_") AS FileNumber, [SumOfAmount]-Nz([Discount]) AS BalanceCalculated, tblCase.*, [TB Time Keeping].[Bill Open], [tblCase].[Last_Name] & ", " & [tblCase].[First_Name] AS Name

FROM tblCase INNER JOIN ([TB Time Keeping] INNER JOIN qry_time_table_totals ON [TB Time Keeping].Bill_ID = qry_time_table_totals.Bill_ID) ON tblCase.CaseID =...
```
### qryTimeTableRunTot
```sql
SELECT qryTTAmount.Time_ID, qryTTAmount.Bill_ID, qryTTAmount.Time_, qryTTAmount.Rate, qryTTAmount.Amount, (Select sum(amount) from qryTTAmount as TT where time_ID<= qryTTAmount.Time_ID and bill_ID=GetBill_ID()) AS Run_total

FROM qryTTAmount

WHERE (((qryTTAmount.Bill_ID)=GetBill_ID()));
```
### qryToBeClosed
```sql
SELECT TblCase.CaseID, TblCase.CaseOpenDate, [Last_Name] & ", " & [First_Name] AS ClientName, TblCase.Case_Letter, TblCase.yr, TblCase.Number_, TblCase.Orig_Atty, Replace([Case_Letter] & [yr] & "-" & [Number_] & "-" & [Orig_Atty],"__","_") AS FileNumber, TblCase.Closed, TblCase.Clsdate, TblCase.ARTrustZero, TblCase.Retainer, TblCase.FamilyLaw

FROM TblCase

WHERE (((TblCase.Closed)=No) AND ((TblCase.FamilyLaw)=Yes))

ORDER BY TblCase.Number_;
```
### qryToBeScanned
```sql
SELECT TblCase.CaseID, TblCase.CaseOpenDate, [Last_Name] & ", " & [First_Name] AS ClientName, TblCase.Case_Letter, TblCase.yr, TblCase.Number_, TblCase.Orig_Atty, Replace([Case_Letter] & [yr] & "-" & [Number_] & "-" & [Orig_Atty],"__","_") AS FileNumber, TblCase.[Scan Location], TblCase.HandlingAtty_Case, TblCase.Closed, TblCase.Scan, TblCase.Clsdate, TblCase.ScanNotAvail

FROM TblCase

WHERE (((TblCase.[Scan Location]) Is Null) AND ((TblCase.Closed)=Yes) AND ((TblCase.ScanNotAvail)=No))

ORDER ...
```
### qryTrustAccount
```sql
SELECT vwTrustAccount.CaseID, vwTrustAccount.SumOfDebit, vwTrustAccount.SumOfCredit, vwTrustAccount.Balance, vwTrustAccount.OrderNr, vwTrustAccount.TrustAccountID

FROM vwTrustAccount

ORDER BY vwTrustAccount.CaseID, vwTrustAccount.OrderNr;
```
### qryTrustAccountBalanceTotals
```sql
SELECT vwTrustAccountBalanceTotals.CaseID, Max(vwTrustAccountBalanceTotals.SumOfBalance) AS SumOfBalance

FROM vwTrustAccountBalanceTotals

GROUP BY vwTrustAccountBalanceTotals.CaseID;
```
### qryTrustAccountBalanceTotals_OLD
```sql
SELECT qryTrustAccount.CaseID, Sum(qryTrustAccount.Balance) AS SumOfBalance

FROM qryTrustAccount

GROUP BY qryTrustAccount.CaseID;
```
### qryTrustAccount_OLD
```sql
SELECT [Trust Account].CaseID, Sum([Trust Account].Debit) AS SumOfDebit, Sum([Trust Account].Credit) AS SumOfCredit, Sum(Nz([debit],0)-Nz([Credit],0)) AS Balance, [Trust Account].OrderNr, [Trust Account].TrustAccountID

FROM [Trust Account]

GROUP BY [Trust Account].CaseID, [Trust Account].OrderNr, [Trust Account].TrustAccountID

ORDER BY [Trust Account].CaseID, [Trust Account].OrderNr;
```
### qryTrustCostsExpended
```sql
SELECT tblCase.CaseID, [Trust Account].TDate, [Trust Account].TMatter, Sum([Trust Account].Credit) AS SumOfCredit, Sum(Nz([Credit],0)) AS CostBalance

FROM tblCase INNER JOIN [Trust Account] ON tblCase.CaseID = [Trust Account].CaseID

GROUP BY tblCase.CaseID, [Trust Account].TDate, [Trust Account].TMatter

HAVING ((([Trust Account].TMatter) Not Like "*Earned*" And ([Trust Account].TMatter) Not Like "*Refund*" And ([Trust Account].TMatter) Not Like "*wire*" And ([Trust Account].TMatter) Not Like ...
```
### qryTrustCostsExpendedTotals
```sql
SELECT vwTrustCostsExpendedTotals.CaseID, vwTrustCostsExpendedTotals.SumOfCostBalance

FROM vwTrustCostsExpendedTotals;
```
### qryTrustCostsExpendedTotals_OLD
```sql
SELECT Sum(qryTrustCostsExpended.CostBalance) AS SumOfCostBalance, qryTrustCostsExpended.CaseID

FROM qryTrustCostsExpended

GROUP BY qryTrustCostsExpended.CaseID;
```
### qryTrustEntriesChron
```sql
SELECT tblCase.CaseID, tblCase.Case_Letter, tblCase.yr, tblCase.Number_, tblCase.Orig_Atty, tblCase.Matter_type, tblCase.Retainer, tblCase.CaseOpenDate, [Trust Account].TrustAccountID, [Trust Account].TDate, [Trust Account].TMatter, [Trust Account].Debit, [Trust Account].Credit, [Trust Account].CheckCashed, [Trust Account].CheckNumber, [case_Letter] & [yr] & "-" & [Number_] & "-" & [Orig_Atty] AS [Case No], [Trust Account].DepCleared, [Trust Account].Reconciled, tblCase.Last_Name, tblCase.First_...
```
### qryTrustEntriesChron65
```sql
SELECT tblCase.CaseID, tblCase.Case_Letter, tblCase.yr, tblCase.Number_, tblCase.Orig_Atty, tblCase.Matter_type, tblCase.Retainer, tblCase.CaseOpenDate, [Trust Account].TrustAccountID, [Trust Account].TDate, [Trust Account].TMatter, [Trust Account].Debit, [Trust Account].Credit, [Trust Account].CheckCashed, [Trust Account].CheckNumber, [case_Letter] & [yr] & "-" & [Number_] & "-" & [Orig_Atty] AS [Case No], [Trust Account].DepCleared, [Trust Account].Reconciled, tblCase.Last_Name, tblCase.First_...
```
### qryTrustEntriesChronRPT35
```sql
SELECT tblCase.CaseID, tblCase.Case_Letter, tblCase.yr, tblCase.Number_, tblCase.Orig_Atty, tblCase.Matter_type, tblCase.Retainer, tblCase.CaseOpenDate, [Trust Account].TrustAccountID, [Trust Account].TDate, [Trust Account].TMatter, [Trust Account].Debit, [Trust Account].Credit, [Trust Account].CheckCashed, [Trust Account].CheckNumber, [case_Letter] & [yr] & "-" & [Number_] & "-" & [Orig_Atty] AS [Case No], [Trust Account].DepCleared, [Trust Account].Reconciled, tblCase.Last_Name, tblCase.First_...
```
### qryTrustEntriesChronRPT35D
```sql
SELECT tblCase.CaseID, tblCase.Case_Letter, tblCase.yr, tblCase.Number_, tblCase.Orig_Atty, tblCase.Matter_type, tblCase.Retainer, tblCase.CaseOpenDate, [Trust Account].TrustAccountID, [Trust Account].TDate, [Trust Account].TMatter, [Trust Account].Debit, [Trust Account].Credit, [Trust Account].CheckCashed, [Trust Account].CheckNumber, [case_Letter] & [yr] & "-" & [Number_] & "-" & [Orig_Atty] AS [Case No], [Trust Account].DepCleared, [Trust Account].Reconciled, tblCase.Last_Name, tblCase.First_...
```
### qryTrustEntriesChronRPT35W
```sql
SELECT tblCase.CaseID, tblCase.Case_Letter, tblCase.yr, tblCase.Number_, tblCase.Orig_Atty, tblCase.Matter_type, tblCase.Retainer, tblCase.CaseOpenDate, [Trust Account].TrustAccountID, [Trust Account].TDate, [Trust Account].TMatter, [Trust Account].Debit, [Trust Account].Credit, [Trust Account].CheckCashed, [Trust Account].CheckNumber, [case_Letter] & [yr] & "-" & [Number_] & "-" & [Orig_Atty] AS [Case No], [Trust Account].DepCleared, [Trust Account].Reconciled, tblCase.Last_Name, tblCase.First_...
```
### qryTrustEntriesChronRPT65D
```sql
SELECT tblCase.CaseID, tblCase.Case_Letter, tblCase.yr, tblCase.Number_, tblCase.Orig_Atty, tblCase.Matter_type, tblCase.Retainer, tblCase.CaseOpenDate, [Trust Account].TrustAccountID, [Trust Account].TDate, [Trust Account].TMatter, [Trust Account].Debit, [Trust Account].Credit, [Trust Account].CheckCashed, [Trust Account].CheckNumber, [case_Letter] & [yr] & "-" & [Number_] & "-" & [Orig_Atty] AS [Case No], [Trust Account].DepCleared, [Trust Account].Reconciled, tblCase.Last_Name, tblCase.First_...
```
### qryTrustEntriesChronRPT65W
```sql
SELECT tblCase.CaseID, tblCase.Case_Letter, tblCase.yr, tblCase.Number_, tblCase.Orig_Atty, tblCase.Matter_type, tblCase.Retainer, tblCase.CaseOpenDate, [Trust Account].TrustAccountID, [Trust Account].TDate, [Trust Account].TMatter, [Trust Account].Debit, [Trust Account].Credit, [Trust Account].CheckCashed, [Trust Account].CheckNumber, [case_Letter] & [yr] & "-" & [Number_] & "-" & [Orig_Atty] AS [Case No], [Trust Account].DepCleared, [Trust Account].Reconciled, tblCase.Last_Name, tblCase.First_...
```
### qryTrustEntriesChronRPT95
```sql
SELECT tblCase.CaseID, tblCase.Case_Letter, tblCase.yr, tblCase.Number_, tblCase.Orig_Atty, tblCase.Matter_type, tblCase.Retainer, tblCase.CaseOpenDate, [Trust Account].TrustAccountID, [Trust Account].TDate, [Trust Account].TMatter, [Trust Account].Debit, [Trust Account].Credit, [Trust Account].CheckCashed, [Trust Account].CheckNumber, [case_Letter] & [yr] & "-" & [Number_] & "-" & [Orig_Atty] AS [Case No], [Trust Account].DepCleared, [Trust Account].Reconciled, tblCase.Last_Name, tblCase.First_...
```
### qryTrustEntriesChronRPT95D
```sql
SELECT tblCase.CaseID, tblCase.Case_Letter, tblCase.yr, tblCase.Number_, tblCase.Orig_Atty, tblCase.Matter_type, tblCase.Retainer, tblCase.CaseOpenDate, [Trust Account].TrustAccountID, [Trust Account].TDate, [Trust Account].TMatter, [Trust Account].Debit, [Trust Account].Credit, [Trust Account].CheckCashed, [Trust Account].CheckNumber, [case_Letter] & [yr] & "-" & [Number_] & "-" & [Orig_Atty] AS [Case No], [Trust Account].DepCleared, [Trust Account].Reconciled, tblCase.Last_Name, tblCase.First_...
```
### qryTrustEntriesChronRPT95W
```sql
SELECT tblCase.CaseID, tblCase.Case_Letter, tblCase.yr, tblCase.Number_, tblCase.Orig_Atty, tblCase.Matter_type, tblCase.Retainer, tblCase.CaseOpenDate, [Trust Account].TrustAccountID, [Trust Account].TDate, [Trust Account].TMatter, [Trust Account].Debit, [Trust Account].Credit, [Trust Account].CheckCashed, [Trust Account].CheckNumber, [case_Letter] & [yr] & "-" & [Number_] & "-" & [Orig_Atty] AS [Case No], [Trust Account].DepCleared, [Trust Account].Reconciled, tblCase.Last_Name, tblCase.First_...
```
### qryTrustReportRPT
```sql
SELECT vwTrustReportRPT.CaseID, vwTrustReportRPT.Case_Letter, vwTrustReportRPT.yr, vwTrustReportRPT.Number_, vwTrustReportRPT.Orig_Atty, vwTrustReportRPT.Matter_type, vwTrustReportRPT.CaseOpenDate, vwTrustReportRPT.CheckCashed, vwTrustReportRPT.CaseNo, vwTrustReportRPT.Last_Name, vwTrustReportRPT.First_Name, vwTrustReportRPT.TrustAccountID

FROM vwTrustReportRPT;
```
### qryTrustReportRPT1
```sql
SELECT vwTrustReportRPT1.CaseID, vwTrustReportRPT1.Case_Letter, vwTrustReportRPT1.yr, vwTrustReportRPT1.Number_, vwTrustReportRPT1.Orig_Atty, vwTrustReportRPT1.Matter_type, vwTrustReportRPT1.CaseOpenDate, vwTrustReportRPT1.SumOfBalance, vwTrustReportRPT1.CheckCashed, vwTrustReportRPT1.CaseNo, vwTrustReportRPT1.Last_Name, vwTrustReportRPT1.First_Name

FROM vwTrustReportRPT1;
```
### qryTrustReportRPT1_OLD
```sql
SELECT qryTrustReportRPT.CaseID, qryTrustReportRPT.Case_Letter, qryTrustReportRPT.yr, qryTrustReportRPT.Number_, qryTrustReportRPT.Orig_Atty, qryTrustReportRPT.Matter_type, qryTrustReportRPT.CaseOpenDate, Sum(qryTrustReportRPT.Balance) AS SumOfBalance, qryTrustReportRPT.CheckCashed, qryTrustReportRPT.CaseNo, qryTrustReportRPT.Last_Name, qryTrustReportRPT.First_Name

FROM (qryTrustReportRPT LEFT JOIN tblCase ON qryTrustReportRPT.CaseID = tblCase.CaseID) LEFT JOIN qryTrustAccount ON qryTrustReport...
```
### qryTrustReportRPT_OLD
```sql
SELECT tblCase.CaseID, tblCase.Case_Letter, tblCase.yr, tblCase.Number_, tblCase.Orig_Atty, tblCase.Matter_type, tblCase.CaseOpenDate, [Trust Account].CheckCashed, [case_Letter] & [yr] & "-" & [Number_] & "-" & [Orig_Atty] AS CaseNo, tblCase.Last_Name, tblCase.First_Name, [Trust Account].TrustAccountID

FROM tblCase LEFT JOIN [Trust Account] ON tblCase.CaseID = [Trust Account].CaseID

WHERE (((tblCase.CaseID) Is Not Null));
```
### qryTrustTotalEarned
```sql
SELECT Sum([Trust Account].Credit) AS SumOfCredit, [Trust Account].TMatter, [Trust Account].CaseID

FROM [Trust Account]

GROUP BY [Trust Account].TMatter, [Trust Account].CaseID

HAVING ((([Trust Account].TMatter) Like "*Earned*"));
```
### qryTrustTotalEarnedSum
```sql
SELECT vwTrustTotalEarnedSum.CaseID, vwTrustTotalEarnedSum.SumOfSumOfCredit

FROM vwTrustTotalEarnedSum;
```
### qryTrustTotalEarnedSum_OLD
```sql
SELECT Sum(qryTrustTotalEarned.SumOfCredit) AS SumOfSumOfCredit, tblCase.CaseID

FROM qryTrustTotalEarned INNER JOIN tblCase ON qryTrustTotalEarned.CaseID = tblCase.CaseID

GROUP BY tblCase.CaseID;
```
### qryUpcomingHearings
```sql
SELECT tblCase.CaseID, tblCase.Orig_Atty, tblCase.HandlingAtty_Case, tblCase.ParaLegal, tblCase.Matter_type, tblHearingDate.Hearing_Date, tblHearingDate.HearingType, [case_Letter] & [yr] & "-" & [Number_] & "-" & [Orig_Atty] AS [Case No], tblCase.Last_Name, tblCase.First_Name, tblCase.Case_Letter, tblCase.yr, tblCase.Number_, [Last_Name] & ", " & [First_Name] AS Name, tblHearingDate.HearingTime, tblHearingDate.HrgCal, tblCase.HmPhone, tblCase.Email, tblHearingDate.Reminder, tblHearingDate.Client...
```
### qryUpdateattyEmail
```sql
UPDATE tblCase SET tblCase.Extended_Ledger = "kbigus@tatebywater.com"

WHERE (((tblCase.Orig_Atty)="KDB"));
```
### qry_CtNames_list_options
```sql
SELECT tblDropD.CodeVal, tblDropD.SortOrder

FROM tblDropD

WHERE (((tblDropD.FieldName)='CtNames'))

ORDER BY tblDropD.SortOrder;
```
### qry_CtType_list_options
```sql
SELECT tblDropD.CodeVal, tblDropD.SortOrder, tblDropD.FieldName

FROM tblDropD

WHERE (((tblDropD.FieldName)="Ctype"))

ORDER BY tblDropD.SortOrder;
```
### qry_FLChildCustodian_list_options
```sql
SELECT tblDropD.CodeVal, tblDropD.SortOrder, tblDropD.FieldName

FROM tblDropD

WHERE (((tblDropD.FieldName)="ChildCustodian"))

ORDER BY tblDropD.SortOrder;
```
### qry_FLCompltMethod_list_options
```sql
SELECT tblDropD.CodeVal, tblDropD.SortOrder, tblDropD.FieldName

FROM tblDropD

WHERE (((tblDropD.FieldName)="CompltMethod"))

ORDER BY tblDropD.SortOrder;
```
### qry_FLDivorceGrounds_list_options
```sql
SELECT tblDropD.CodeVal, tblDropD.FieldName, tblDropD.SortOrder

FROM tblDropD

WHERE (((tblDropD.FieldName)="DivorceGrounds"))

ORDER BY tblDropD.SortOrder;
```
### qry_FLLengthSeparation_list_options
```sql
SELECT tblDropD.CodeVal, tblDropD.FieldName, tblDropD.SortOrder

FROM tblDropD

WHERE (((tblDropD.FieldName)="LengthSeparation"))

ORDER BY tblDropD.SortOrder;
```
### qry_FLNOHMethod_list_options
```sql
SELECT tblDropD.CodeVal, tblDropD.SortOrder, tblDropD.FieldName

FROM tblDropD

WHERE (((tblDropD.FieldName)="NOHMethod"))

ORDER BY tblDropD.SortOrder;
```
### qry_FLNumberChildren_list_options
```sql
SELECT tblDropD.CodeVal, tblDropD.SortOrder, tblDropD.FieldName

FROM tblDropD

WHERE (((tblDropD.FieldName)="NumberChildren"))

ORDER BY tblDropD.SortOrder;
```
### qry_HearingType_list_options
```sql
SELECT tblDropD.CodeVal, tblDropD.FieldName, tblDropD.SortOrder

FROM tblDropD

WHERE (((tblDropD.FieldName)="Hearingtype"))

ORDER BY tblDropD.SortOrder;
```
### qry_InvoiceAR_curr
```sql
SELECT qry_current_invoice.Date2, qry_current_invoice.Pay_Outlay, qry_current_invoice.Charge, qry_current_invoice.OrderNr, qry_current_invoice.CaseID

FROM qry_current_invoice

WHERE (((qry_current_invoice.Charge)>0));
```
### qry_InvoicePymts_curr
```sql
SELECT qry_current_invoice.CaseID, qry_current_invoice.OrderNr, qry_current_invoice.Payment, qry_current_invoice.Pay_Outlay, qry_current_invoice.Date2

FROM qry_current_invoice

WHERE (((qry_current_invoice.Payment)<>0));
```
### qry_LastINV
```sql
SELECT tbl_InvoiceSent.CaseID, Last(tbl_InvoiceSent.InvSent) AS LastOfInvSent

FROM tbl_InvoiceSent

GROUP BY tbl_InvoiceSent.CaseID;
```
### qry_OrigAtty_list_options
```sql
SELECT tblDropD.Code, tblDropD.CodeVal, tblDropD.FieldName, tblDropD.SortOrder

FROM tblDropD

WHERE (((tblDropD.FieldName)="Orig_Atty"))

ORDER BY tblDropD.SortOrder;
```
### qry_RetBalSums_by_PastDue
```sql
SELECT qryOutstandingARRPT1.[Past Due], qryOutstandingARRPT1.Orig_Atty, Sum(qryOutstandingARRPT1.RetBal) AS SumOfRetBal

FROM qryOutstandingARRPT1

GROUP BY qryOutstandingARRPT1.[Past Due], qryOutstandingARRPT1.Orig_Atty;
```
### qry_TA_uncashed_checks
```sql
SELECT [Last_Name] & ", " & [First_Name] AS Name, Replace([Case_Letter] & [yr] & "-" & [Number_] & "-" & [Orig_Atty],"__","_") AS FileNumber, [Trust Account].*

FROM TblCase INNER JOIN [Trust Account] ON TblCase.CaseID = [Trust Account].CaseID

WHERE ((([Trust Account].CheckCashed)=No))

ORDER BY [Last_Name] & ", " & [First_Name];
```
### qry_TimeKeeping_CaseID_totals
```sql
SELECT qry_TimeKeeping_bill_totals.CaseID, Sum(qry_TimeKeeping_bill_totals.Total) AS SumOfTotal, qry_TimeKeeping_bill_totals.[Bill Closed], qry_TimeKeeping_bill_totals.IANumber, qry_TimeKeeping_bill_totals.Bill_ID

FROM qry_TimeKeeping_bill_totals

GROUP BY qry_TimeKeeping_bill_totals.CaseID, qry_TimeKeeping_bill_totals.[Bill Closed], qry_TimeKeeping_bill_totals.IANumber, qry_TimeKeeping_bill_totals.Bill_ID;
```
### qry_TimeKeeping_bill_totals
```sql
SELECT [TB Time Keeping].CaseID, tblTimeTableDetail.Bill_ID, tblTimeTableDetail.Rate, tblTimeTableDetail.Time_, [Rate]*[Time_] AS Total, [TB Time Keeping].[Bill Closed], [TB Time Keeping].IANumber

FROM [TB Time Keeping] INNER JOIN tblTimeTableDetail ON [TB Time Keeping].Bill_ID = tblTimeTableDetail.Bill_ID

WHERE ((([TB Time Keeping].[Bill Closed])=No));
```
### qry_advanced_nonadvanced_payments
```sql
SELECT [Matter and AR].MatterID, IIf([charge]>0 And [firmprepaid]=False,[charge],0) AS NonAdvancedCharges, IIf([charge]>0 And [firmprepaid]=True,[charge],0) AS AdvancedCharges, IIf([payment]>0 And [firmprepaid]=0,[payment],0) AS PaymentMade, [NonAdvancedCharges]+[AdvancedCharges]-[PaymentMade] AS PreBalance, [Matter and AR].CaseID

FROM tblCase INNER JOIN [Matter and AR] ON tblCase.CaseID = [Matter and AR].CaseID

ORDER BY [Matter and AR].OrderNr;
```
### qry_advanced_payments
```sql
SELECT p.Name, p.FileNumber, p.MatterID, p.CaseID, p.Date2, p.Pay_Outlay, p.Charge, p.Payment, p.FirmPrepaid, p.InsertPymt, p.AdvancedLegal, p.SSMA_TimeStamp, p.Orig_Atty, p.Case_Letter, p.CodeVal, p.Creimb, CCur(Nz (t.SumOfBalance_agg, 0)) AS SumOfBalance

FROM vw_advanced_payments AS p LEFT JOIN (SELECT
            CaseID,
            Max(SumOfBalance) AS SumOfBalance_agg
        FROM
            qryTrustAccountBalanceTotals
        GROUP BY
            CaseID
    )  AS t ON p.CaseID = t.CaseI...
```
### qry_advanced_payments_OLD
```sql
SELECT [Last_Name] & ", " & [First_Name] AS Name, Replace([Case_Letter] & [yr] & "-" & [Number_] & "-" & [Orig_Atty],"__","_") AS FileNumber, [Matter and AR].*, TblCase.Orig_Atty, TblCase.Case_Letter, tblDropD.CodeVal, [Matter and AR].Date2

FROM (TblCase LEFT JOIN tblDropD ON TblCase.Case_Letter = tblDropD.Code) INNER JOIN [Matter and AR] ON TblCase.CaseID = [Matter and AR].CaseID

WHERE ((([Matter and AR].FirmPrepaid)=Yes))

ORDER BY [Matter and AR].Date2 DESC;
```
### qry_advanced_totals
```sql
SELECT tblCase.CaseID, [Matter and AR].Charge, [Matter and AR].FirmPrepaid

FROM tblCase INNER JOIN [Matter and AR] ON tblCase.CaseID = [Matter and AR].CaseID

WHERE ((([Matter and AR].FirmPrepaid)=Yes));
```
### qry_advanced_totals_SUM
```sql
SELECT vw_advanced_totals_SUM.CaseID, vw_advanced_totals_SUM.SumOfCharge, vw_advanced_totals_SUM.FirmPrepaid

FROM vw_advanced_totals_SUM;
```
### qry_advanced_totals_SUM_OLD
```sql
SELECT qry_advanced_totals.CaseID, Sum(qry_advanced_totals.Charge) AS SumOfCharge, qry_advanced_totals.FirmPrepaid

FROM qry_advanced_totals

GROUP BY qry_advanced_totals.CaseID, qry_advanced_totals.FirmPrepaid;
```
### qry_caseID_clients
```sql
SELECT [last_name] & ", " & [first_name] AS Name, TblCase.CaseID, TblCase.Closed

FROM TblCase

GROUP BY [last_name] & ", " & [first_name], TblCase.CaseID, TblCase.Closed, TblCase.Last_Name

HAVING (((TblCase.Closed)=No) AND ((TblCase.Last_Name) Is Not Null))

ORDER BY [last_name] & ", " & [first_name];
```
### qry_client_names
```sql
SELECT tblCase.CaseID, [Last_Name] & ", " & [First_Name] & " " & Replace([Case_Letter] & [Yr] & "-" & [Number_] & "-" & [Orig_Atty],"__","_") AS Name

FROM tblCase

ORDER BY [Last_Name] & ", " & [First_Name] & " " & Replace([Case_Letter] & [Yr] & "-" & [Number_] & "-" & [Orig_Atty],"__","_");
```
### qry_client_names_TK
```sql
SELECT TblCase.CaseID, [Last_Name] & ", " & [First_Name] & " " & Replace([Case_Letter] & [Yr] & "-" & [Number_] & "-" & [Orig_Atty],"__","_") AS Name, TblCase.Closed

FROM TblCase

GROUP BY TblCase.CaseID, [Last_Name] & ", " & [First_Name] & " " & Replace([Case_Letter] & [Yr] & "-" & [Number_] & "-" & [Orig_Atty],"__","_"), TblCase.Closed, TblCase.Last_Name

HAVING (((TblCase.Closed)=No) AND ((TblCase.Last_Name) Is Not Null))

ORDER BY [Last_Name] & ", " & [First_Name] & " " & Replace([Case_Lett...
```
### qry_current_invoice
```sql
SELECT vw_current_invoice.Balance, vw_current_invoice.[Balance Due Date], vw_current_invoice.[Billing Notes], vw_current_invoice.Last_Name, vw_current_invoice.First_Name, vw_current_invoice.CaseOpenDate, vw_current_invoice.yr, vw_current_invoice.Number_, vw_current_invoice.Orig_Atty, vw_current_invoice.Address, vw_current_invoice.City, vw_current_invoice.State, vw_current_invoice.Zip, vw_current_invoice.Matter_type, vw_current_invoice.Retainer, vw_current_invoice.Case_Letter, vw_current_invoice....
```
### qry_current_invoice_OLD
```sql
SELECT fncGetMatterARBalanceWithCaseID([OrderNr],[tblcase].[CaseID]) AS Balance, Billing.[Balance Due Date], Billing.[Billing Notes], tblCase.Last_Name, tblCase.First_Name, tblCase.CaseOpenDate, tblCase.yr, tblCase.Number_, tblCase.Orig_Atty, tblCase.Address, tblCase.City, tblCase.State, tblCase.Zip, tblCase.Matter_type, tblCase.Retainer, tblCase.Case_Letter, [Matter and AR].OrderNr, tblCase.Retainer, [Matter and AR].CaseID, [Matter and AR].Date2, [Matter and AR].Pay_Outlay, [Matter and AR].Char...
```
### qry_disposition_closingSheet
```sql
SELECT TblCase.CaseID, TblCase.Last_Name, TblCase.First_Name, TblCase.Address, TblCase.City, TblCase.State, TblCase.Zip, TblCase.DOB, TblCase.SSN, TblCase.HmPhone, TblCase.OtherPhone, TblCase.Email, TblCase.Comments, TblCase.Referral, TblCase.Closed, TblCase.Clsdate, [tblCase].[case_Letter] & [tblCase].[yr] & "-" & [Number_] & "-" & [tblCase].[Orig_Atty] AS CaseNo, Disposition.Dispo_Date, Disposition.DispoJudge, Disposition.Dispo_Atty, Disposition.DispoOppC, Disposition.[PI Settlement Amount], D...
```
### qry_file_numbers
```sql
SELECT tblCase.CaseID, Replace([Case_Letter] & [Yr] & "-" & [Number_] & "-" & [Orig_Atty],"__","_") AS FileNo

FROM tblCase;
```
### qry_find_table_by_field_name
```sql
SELECT tblFields.*, tblFields.FieldName

FROM tblFields

WHERE (((tblFields.FieldName)="CostHold"));
```
### qry_get_MatterID_from_zero_balance
```sql
SELECT qryNewInvoice_01.CaseID, qryNewInvoice_01.MatterID, qryNewInvoice_01.RetBal

FROM qryNewInvoice_01;
```
### qry_get_time_keeping_numbers
```sql
SELECT tblCase.CaseID, Count([TB Time Keeping].IANumber) AS CountOfIANumber

FROM tblCase INNER JOIN [TB Time Keeping] ON tblCase.CaseID = [TB Time Keeping].CaseID

GROUP BY tblCase.CaseID

HAVING (((Count([TB Time Keeping].IANumber)) Is Not Null));
```
### qry_invoice_comprehensive_trust_acc_cur
```sql
SELECT vw_invoice_comprehensive_trust_acc_cur_unfiltered.CaseID, vw_invoice_comprehensive_trust_acc_cur_unfiltered.OrderNr, vw_invoice_comprehensive_trust_acc_cur_unfiltered.TDate, vw_invoice_comprehensive_trust_acc_cur_unfiltered.TMatter, vw_invoice_comprehensive_trust_acc_cur_unfiltered.Debit, vw_invoice_comprehensive_trust_acc_cur_unfiltered.balance

FROM vw_invoice_comprehensive_trust_acc_cur_unfiltered

WHERE (((vw_invoice_comprehensive_trust_acc_cur_unfiltered.TMatter) Not Like "*Earned*" ...
```
### qry_invoice_comprehensive_trust_acc_cur_OLD
```sql
SELECT [Trust Account].CaseID, [Trust Account].OrderNr, [Trust Account].TDate, [Trust Account].TMatter, [Trust Account].Debit, fncGetTABalanceWithCaseID([ordernr],[caseid]) AS Balance

FROM [Trust Account]

WHERE ((([Trust Account].TMatter) Not Like "*Earned*" And ([Trust Account].TMatter) Not Like "*Reimb*" And ([Trust Account].TMatter) Not Like "*Wire*" And ([Trust Account].TMatter) Not Like "*refund*"))

ORDER BY [Trust Account].OrderNr;
```
### qry_invoice_comprehensive_trust_acc_cur_unfiltered
```sql
SELECT vw_invoice_comprehensive_trust_acc_cur_unfiltered.CaseID, vw_invoice_comprehensive_trust_acc_cur_unfiltered.OrderNr, vw_invoice_comprehensive_trust_acc_cur_unfiltered.TDate, vw_invoice_comprehensive_trust_acc_cur_unfiltered.TMatter, vw_invoice_comprehensive_trust_acc_cur_unfiltered.Debit, vw_invoice_comprehensive_trust_acc_cur_unfiltered.balance

FROM vw_invoice_comprehensive_trust_acc_cur_unfiltered

ORDER BY vw_invoice_comprehensive_trust_acc_cur_unfiltered.OrderNr;
```
### qry_invoice_comprehensive_trust_acc_cur_unfiltered_old
```sql
SELECT [Trust Account].CaseID, [Trust Account].OrderNr, [Trust Account].TDate, [Trust Account].TMatter, [Trust Account].Debit, fncGetTABalanceWithCaseID([ordernr],[caseid]) AS Balance

FROM [Trust Account]

ORDER BY [Trust Account].OrderNr;
```
### qry_invoices_summary
```sql
SELECT vw_invoices_summary.CaseID, vw_invoices_summary.Name, vw_invoices_summary.First_Name, vw_invoices_summary.Last_Name, vw_invoices_summary.Retainer, vw_invoices_summary.SumOfCharge, vw_invoices_summary.SumOfPayment, vw_invoices_summary.SumOfBalance, vw_invoices_summary.BalanceCalculated, vw_invoices_summary.BalRetCalculated, vw_invoices_summary.FileNumber, vw_invoices_summary.[Balance Due Date], vw_invoices_summary.Orig_Atty, vw_invoices_summary.HandlingAtty_Case, vw_invoices_summary.CodeVa...
```
### qry_invoices_summaryRPT
```sql
SELECT qryInvoiceRPT1.CaseID, [Last_Name] & ", " & [First_Name] AS Name, qryInvoiceRPT1.First_Name, qryInvoiceRPT1.Last_Name, qryInvoiceRPT1.Retainer, Sum(qryInvoiceRPT1.Charge) AS SumOfCharge, Sum(qryInvoiceRPT1.Payment) AS SumOfPayment, Sum(qryInvoiceRPT1.Balance) AS SumOfBalance, Sum([charge]-[payment]) AS BalanceCalculated, [BalanceCalculated]+[qryInvoiceRPT1].[Retainer] AS BalRetCalculated, qryInvoiceRPT1.Spanish, Replace([Case_Letter] & [yr] & "-" & [Number_] & "-" & [Orig_Atty],"__","_") ...
```
### qry_invoices_summary_OLD
```sql
SELECT qryInvoiceRPT1.CaseID, [Last_Name] & ", " & [First_Name] AS Name, qryInvoiceRPT1.First_Name, qryInvoiceRPT1.Last_Name, qryInvoiceRPT1.Retainer, Sum(qryInvoiceRPT1.Charge) AS SumOfCharge, Sum(qryInvoiceRPT1.Payment) AS SumOfPayment, qryInvoiceRPT1.Balance, Sum([charge]-[payment]) AS BalanceCalculated, [BalanceCalculated]+[qryInvoiceRPT1].[Retainer] AS BalRetCalculated, [Balance]+[qryInvoiceRPT1].[Retainer] AS BalRet, qryInvoiceRPT1.[Past Due], qryInvoiceRPT1.chkBalanceDue, Replace([Case_Le...
```
### qry_last_invoice_sent
```sql
SELECT vw_last_invoice_sent.CaseID, vw_last_invoice_sent.LastOfInvSent

FROM vw_last_invoice_sent

GROUP BY vw_last_invoice_sent.CaseID, vw_last_invoice_sent.LastOfInvSent;
```
### qry_last_invoice_sent_OLD
```sql
SELECT qry_invoices_summary.CaseID, Last(tbl_InvoiceSent.InvSent) AS LastOfInvSent

FROM qry_invoices_summary INNER JOIN tbl_InvoiceSent ON qry_invoices_summary.CaseID = tbl_InvoiceSent.CaseID

GROUP BY qry_invoices_summary.CaseID, tbl_InvoiceSent.[TK Sent]

HAVING (((tbl_InvoiceSent.[TK Sent])=No));
```
### qry_matterAR_pay_putlay_list_options
```sql
SELECT tblDropD.codeVal, tblDropD.sortOrder

FROM tblDropD

WHERE (((tblDropD.FieldName)='MatterAR_Pay_Outlay'))

ORDER BY tblDropD.sortOrder;
```
### qry_max_matterID_by_orderNr
```sql
SELECT vw_max_matterID_by_orderNr.MaxOfOrderNr, vw_max_matterID_by_orderNr.MaxOfMatterID, vw_max_matterID_by_orderNr.CaseID

FROM vw_max_matterID_by_orderNr;
```
### qry_max_matterID_by_orderNr_OLD
```sql
SELECT Max([Matter and AR].OrderNr) AS MaxOfOrderNr, Max([Matter and AR].MatterID) AS MaxOfMatterID, [Matter and AR].CaseID

FROM [Matter and AR]

GROUP BY [Matter and AR].CaseID;
```
### qry_orig_atty
```sql
SELECT DISTINCT tblCase.Orig_Atty

FROM tblCase

GROUP BY tblCase.Orig_Atty;
```
### qry_orig_atty_filter
```sql
SELECT DISTINCT tblCase.CaseID, tblCase.Orig_Atty

FROM tblCase

GROUP BY tblCase.CaseID, tblCase.Orig_Atty;
```
### qry_takeOff_year_month
```sql
SELECT tblTakeOffMonth.TakeOffMonthID, tblTakeOffMonth.TakeOffDate, Year([TakeOffDate]) AS YearOnly, Month([TakeOffDate]) AS MonthOnly

FROM tblTakeOffMonth;
```
### qry_take_off_step2_attorney_sums
```sql
SELECT tblTakeOffMonth.TakeOffMonthID, qryTakeOffStep2.Orig_Atty, Sum(qryTakeOffStep2.EarlyEarned) AS SumOfEarlyEarned, Sum(qryTakeOffStep2.TOEarned) AS SumOfTOEarned, Sum(qryTakeOffStep2.CostReimb) AS SumOfCostReimb, Sum(qryTakeOffStep2.CBHRev) AS SumOfCBHRev, Sum(qryTakeOffStep2.MKRev) AS SumOfMKRev, Sum(qryTakeOffStep2.CBHCom) AS SumOfCBHCom, Sum(qryTakeOffStep2.MTRev) AS SumOfMTRev, Sum(qryTakeOffStep2.MTCom) AS SumOfMTCom, Sum(qryTakeOffStep2.KBCom) AS SumOfKBCom, Sum(qryTakeOffStep2.MKCom)...
```
### qry_take_off_step2_sums
```sql
SELECT vw_take_off_step2_sums.TakeOffMonthID, vw_take_off_step2_sums.SumOfCBHRev, vw_take_off_step2_sums.SumOfMKRev, vw_take_off_step2_sums.SumOfCBHCom, vw_take_off_step2_sums.SumOfMTRev, vw_take_off_step2_sums.SumOfMTCom, vw_take_off_step2_sums.SumOfKBCom, vw_take_off_step2_sums.SumOfMKCom, vw_take_off_step2_sums.SumOfRLFCom, vw_take_off_step2_sums.SumOfEarlyEarned, vw_take_off_step2_sums.SumOfTOEarned, vw_take_off_step2_sums.SumOfTOEarlyAndEarned, vw_take_off_step2_sums.SumOfCostReimb

FROM vw...
```
### qry_take_off_step2_sums_OLD
```sql
SELECT tblTakeOffMonth.TakeOffMonthID, Sum(qryTakeOffStep2.CBHRev) AS SumOfCBHRev, Sum(qryTakeOffStep2.MKRev) AS SumOfMKRev, Sum(qryTakeOffStep2.CBHCom) AS SumOfCBHCom, Sum(qryTakeOffStep2.MTRev) AS SumOfMTRev, Sum(qryTakeOffStep2.MTCom) AS SumOfMTCom, Sum(qryTakeOffStep2.KBCom) AS SumOfKBCom, Sum(qryTakeOffStep2.MKCom) AS SumOfMKCom, Sum(qryTakeOffStep2.RLFCom) AS SumOfRLFCom

FROM qryTakeOffStep2 INNER JOIN tblTakeOffMonth ON qryTakeOffStep2.TakeOffMonthID = tblTakeOffMonth.TakeOffMonthID

GRO...
```
### qry_tblUsers
```sql
SELECT tblUsers.*, tblAccessType.*

FROM tblUsers INNER JOIN tblAccessType ON tblUsers.Access = tblAccessType.AccessType;
```
### qry_time_table_totals
```sql
SELECT qryTTAmount.Bill_ID, Sum(qryTTAmount.Amount) AS SumOfAmount

FROM qryTTAmount

GROUP BY qryTTAmount.Bill_ID;
```
### qry_time_table_totals_SUM
```sql
SELECT vw_time_table_totals_SUM.CaseID, vw_time_table_totals_SUM.SumOfSumOfAmount

FROM vw_time_table_totals_SUM;
```
### qry_time_table_totals_SUM_OLD
```sql
SELECT Sum(qry_time_table_totals_atty.SumOfAmount) AS SumOfSumOfAmount, [TB Time Keeping].CaseID

FROM qry_time_table_totals_atty INNER JOIN [TB Time Keeping] ON qry_time_table_totals_atty.Bill_ID = [TB Time Keeping].Bill_ID

GROUP BY [TB Time Keeping].CaseID;
```
### qry_time_table_totals_atty
```sql
SELECT qryTTAmountAtty.Bill_ID, qryTTAmountAtty.Tatty, Sum(qryTTAmountAtty.Amount) AS SumOfAmount

FROM qryTTAmountAtty

GROUP BY qryTTAmountAtty.Bill_ID, qryTTAmountAtty.Tatty;
```
### qry_time_table_totals_atty_SUM
```sql
SELECT vw_time_table_totals_atty_SUM.CaseID, vw_time_table_totals_atty_SUM.Tatty, vw_time_table_totals_atty_SUM.SumOfSumOfAmount

FROM vw_time_table_totals_atty_SUM;
```
### qry_time_table_totals_atty_SUM_OLD
```sql
SELECT qry_time_table_totals_atty.Tatty, Sum(qry_time_table_totals_atty.SumOfAmount) AS SumOfSumOfAmount, [TB Time Keeping].CaseID

FROM qry_time_table_totals_atty INNER JOIN [TB Time Keeping] ON qry_time_table_totals_atty.Bill_ID = [TB Time Keeping].Bill_ID

GROUP BY qry_time_table_totals_atty.Tatty, [TB Time Keeping].CaseID;
```
### qry_time_table_totals_hours
```sql
SELECT qryTTAmount.Bill_ID, Sum(qryTTAmount.Amount) AS SumOfAmount

FROM qryTTAmount

GROUP BY qryTTAmount.Bill_ID;
```
### qry_time_table_totals_hours_sum
```sql
SELECT qryTTAmount.Bill_ID, Sum(qryTTAmount.Amount) AS SumOfAmount

FROM qryTTAmount

GROUP BY qryTTAmount.Bill_ID;
```
### qry_tmatter_list_options
```sql
SELECT tblDropD.codeVal, tblDropD.sortOrder

FROM tblDropD

WHERE (((tblDropD.FieldName)='TMatter'))

ORDER BY tblDropD.sortOrder;
```
### qry_trustStatements
```sql
SELECT [Last_Name] & ", " & [First_Name] AS Name, tblCase.Matter_type, tblCase.Orig_Atty, Replace([Case_Letter] & [yr] & "-" & [Number_] & "-" & [Orig_Atty],"__","_") AS FileNumber, tblCase.Case_Letter, tblCase.Number_, tblCase.CaseID, tblCase.Closed, qryTakeOff_trust_account.Balance, tblCase.Executor, tblCase.Case_Letter

FROM tblCase LEFT JOIN qryTakeOff_trust_account ON tblCase.CaseID = qryTakeOff_trust_account.CaseID

GROUP BY [Last_Name] & ", " & [First_Name], tblCase.Matter_type, tblCase.O...
```
### qry_uncashed_trust_checks
```sql
SELECT [Last_Name] & ", " & [First_Name] AS Name, Replace([Case_Letter] & [yr] & "-" & [Number_] & "-" & [Orig_Atty],"__","_") AS FileNumber, [Trust Account].CheckCashed, [Trust Account].DepCleared, [Trust Account].CaseID, [Trust Account].TDate, [Trust Account].TMatter, [Trust Account].Debit, [Trust Account].Credit, [Trust Account].CheckNumber

FROM TblCase INNER JOIN [Trust Account] ON TblCase.CaseID = [Trust Account].CaseID

WHERE ((([Trust Account].CheckCashed)=Yes)) OR ((([Trust Account].Dep...
```

## VBA Object Inventory
| Type | Count | Objects |
|------|-------|--------|
| Forms | 94 | Intakes, Time Keeping, frmActionNeeded, frmActionNeededAll, frmActionNeededAll2, frmActionNeededAll3, frmAddUser, frmAdminLoginTK, frmApplicationLoad, frmAttyFeeGeneration, frmAttyNotes, frmBankruptcy, frmBilling, frmBrowse, frmBrowse_BackEnd, frmCalendarCheck, frmCalls, frmCallsList, frmCaseList, frmCaseListAll, frmCaseListClosed, frmCaseListOpen, frmCaseListOpen subform, frmChild, frmClientLedger, frmClientReviews, frmClientsConflict, frmConflictChk, frmCrimStatusReport, frmCtCaseNumbers, frmDisposition, frmDispositions, frmFamilyLaw, frmHearingDate, frmHome, frmHomeAdmin, frmHomeAdminLogin, frmIntakesConflicts, frmInvoiceSent, frmLogin, frmMatter, frmOkAlert, frmOpenReport, frmOppPartyConflict, frmPersInjDemand, frmPersInjLog, frmPersInjLog2, frmPersInjProvider, frmPersInjuryStatusReport, frmPersonalInjury, frmPersonalInjury2, frmReceipt, frmScanLocation, frmScansubform, frmSourceAnalytics, frmSubCH13Plans, frmSubPrevBankrupt, frmSubProofOfClaims, frmTKClose, frmTRUSTENTRIESCHRON, frmTakeOff, frmTakeOff2, frmTakeOffReconciliation, frmTakeOffSteps, frmTakeOffSubForm, frmTakeOffSubForm2, frmTakeOffSubForm3, frmTakeOffSubForm_OLD, frmTakeOffTest, frmTakeOffTotalFeesCosts, frmTimeKeepingClosed, frmTimeKeepingOpen, frmTimeTableDetail, frmTimeTableDetailMerge, frmToBeClosed, frmToBeScanned, frmTrustAccount, frmUpcoming Hearings, frmUsers, frmUsers_Edit, frmYearWiseCaseList, frmYesNoAlert, frm_Billing_Tracker, frm_Billing_Tracker2, frm_advanced_payments, frm_invoices_summary, frm_trust_summary, frm_uncashed_trust_checks, zClient Ledger OLD, zfrmFamilyLaw OLD, zfrmPersInjSOL, zfrmPersonalDetailsFamilyLaw, zfrmSelectCaseNum, zfrmSelectCaseNum_Discount |
| Reports | 99 | Accounts Receivable, Case Sources and Revenue, Client Closing Sheet, Client_Trust_Accounts_for_PreTake_Off, Client_Trust_Accounts_for_Take_Off, Copy Of Client Closing Sheet, Invoice, Invoice - No Balance Due, Invoice - Past Due, Invoice Attach - Hourly, Invoice Attach - Hourly w Discount, Invoice2, New Invoice, Rpt_MergeInvTK, Statement of Trust Account, rptBillingTotals, rptClientNotes, rptComprehensiveTKStatement, rptCriminalStatus, rptCriminalStatusActionNeeded, rptCriminalStatusChargeNos, rptCriminalStatusNotesLog, rptCriminalStatusNotesLog2, rptCriminalStatusUpcHrgs, rptInvoiceComprARCur, rptInvoiceComprPymtsAR, rptInvoiceComprPymtsARCur, rptInvoiceComprTrustCur, rptInvoiceComprehensiveAR, rptInvoiceComprehensiveAR2, rptInvoiceComprehensiveTrust, rptInvoiceComprehensiveTrust2, rptLastTenOpen, rptLastWeekIntake, rptPISOLList, rptPIStatusSOL, rptPersInjProviderBills, rptPersInjStatusAction, rptPersInjStatusDemand, rptPersInjStatusLog, rptPersInjuryStatus, rptReceipt, rptReceiptC, rptReceiptR, rptReceiptRec, rptReconciliation, rptTKReport, rptTKReport2, rpt_Billing_Closing, rpt_CaseNumber_Closing, rpt_Compr_InvoiceADVCur, rpt_Compr_InvoiceStmtCur, rpt_Compr_InvoiceTKExCur, rpt_Comprehensive_Invoice, rpt_Comprehensive_Invoice2, rpt_Comprehensive_InvoiceADV, rpt_Comprehensive_InvoiceADVS, rpt_Comprehensive_InvoiceStmt, rpt_Comprehensive_InvoiceStmtS, rpt_Comprehensive_InvoiceTKEx, rpt_Comprehensive_InvoiceTKEx1, rpt_Comprehensive_InvoiceTKEx1S, rpt_Comprehensive_InvoiceTKEx2, rpt_Comprehensive_InvoiceTKEx2S, rpt_Comprehensive_InvoiceTKEx3Costs, rpt_Comprehensive_InvoiceTKEx3CostsS, rpt_Comprehensive_InvoiceTKLessTrustCostAR, rpt_Comprehensive_InvoiceTKLessTrustRep, rpt_Comprehensive_InvoiceTKLessTrustRep2, rpt_Disposition_Closing, rpt_File_Folder_Label, rpt_Main_Closing, rpt_Matter_Closing, rpt_MergeInvMatter, rpt_MergeInvTimeDetail, rpt_OpenCases, rpt_Open_Cases, rpt_Reconciliation sub, rpt_TKExceedsTrust, rpt_TKLessTrust, rpt_TKTotalAdvance, rpt_TimeDetail_Comprehensive, rpt_TimeDetail_Comprehensive2, rpt_Trust_Chron_35, rpt_Trust_Chron_35D, rpt_Trust_Chron_35W, rpt_Trust_Chron_65, rpt_Trust_Chron_65D, rpt_Trust_Chron_65W, rpt_Trust_Chron_95, rpt_Trust_Chron_95D, rpt_Trust_Chron_95W, rpt_Trust_Closing, rpt_address_label, rpt_address_labelEx, rpt_adj_address_label, rpt_ftrustee_address_label, rpt_opp_counsel_address_label, rpt_trustee_address_label |
| Modules | 23 | AccessType, Authentication, CaseGeneratorModule, Configuration, Context, DocumentManagement, FormUtils, ModGeneric Func, ModUpload, Module1, OutlookApp, PcaStdLib, Relinking, UI Images, User, Util, ValidatedForm, basFindField, clsFormValidation, modErrmsgs, modFutureDateVarification, modGaz, mod_create_table_with_all_db_schema |
| Classes | 0 |  |
| Macros | 0 |  |

## Structured Report Inventory
**Extracted:** 99 reports

| Report | Data Source | Sections | Subreports |
|--------|------------|----------|------------|
| rpt_TKTotalAdvance | qryInvoiceAttachRPT1 | 5 | 0 |
| rpt_Matter_Closing | SELECT vw_rpt_Matter_Closing.CaseID, vw_rpt_Matter_Closing.M... | 3 | 0 |
| rptCriminalStatusNotesLog | SELECT tblNotes.CaseID, tblNotes.NoteDate, tblNotes.NotePers... | 2 | 0 |
| rpt_Trust_Closing | qryStmtTrustRPT1 | 3 | 0 |
| rpt_Billing_Closing | SELECT tblCase.CaseID, Disposition.[Total Earned Fee] AS Exp... | 1 | 0 |
| Accounts Receivable | qry_invoices_summaryRPT | 5 | 0 |
| rpt_Main_Closing | tblCase | 3 | 5 |
| rpt_Disposition_Closing | qry_Disposition_ClosingSheet | 1 | 0 |
| rpt_CaseNumber_Closing | tbl_CtCaseNumbers | 2 | 0 |
| Case Sources and Revenue | qryCaseSourcesRPT1 | 5 | 0 |
| rpt_Comprehensive_InvoiceStmtS | SELECT [TB Time Keeping].[Bill Closed], [TB Time Keeping].IA... | 4 | 4 |
| rpt_Comprehensive_InvoiceTKEx3Costs | SELECT [TB Time Keeping].[Bill Closed], [TB Time Keeping].IA... | 4 | 4 |
| Copy Of Client Closing Sheet | qryClosing RPT1 | 4 | 0 |
| rpt_Compr_InvoiceADVCur | SELECT [TB Time Keeping].[Bill Closed], [TB Time Keeping].IA... | 5 | 4 |
| Client Closing Sheet | qryClosing RPT1 | 4 | 0 |
| rptInvoiceComprehensiveTrust2 | qryInvoiceComprehensiveTrustCredit4 | 3 | 0 |
| Client_Trust_Accounts_for_PreTake_Off | qryTakeOff | 4 | 0 |
| Client_Trust_Accounts_for_Take_Off | qryAttyTrustAcctsTOff | 4 | 0 |
| rpt_Reconciliation sub | SELECT qryTakeOffStep2.FileNumber, qryTakeOffStep2.Name, qry... | 2 | 0 |
| rpt_Comprehensive_Invoice | SELECT [TB Time Keeping].[Bill Closed], [TB Time Keeping].IA... | 4 | 3 |
| New Invoice | qry_current_invoice | 6 | 0 |
| Invoice | qryInvoiceRPT1 | 6 | 0 |
| Invoice - No Balance Due | qryInvoiceRPT1 | 7 | 0 |
| Invoice - Past Due | qryInvoiceRPT1 | 7 | 0 |
| Invoice Attach - Hourly | qryInvoiceAttachRPT1 | 7 | 0 |
| rptPISOLList |  | 5 | 0 |
| rpt_Comprehensive_InvoiceTKEx3CostsS | SELECT [TB Time Keeping].[Bill Closed], [TB Time Keeping].IA... | 5 | 4 |
| Invoice Attach - Hourly w Discount | qryInvoiceAttachRPT1 | 5 | 0 |
| rpt_Trust_Chron_35 | qryTrustEntriesChronRPT35 | 5 | 0 |
| rptInvoiceComprehensiveAR2 | SELECT tblCase.CaseID, [Matter and AR].MatterID, [Matter and... | 3 | 0 |
| Invoice2 | qryInvoiceRPT1 | 7 | 0 |
| rpt_address_label | tblCase | 1 | 0 |
| rpt_address_labelEx | tblCase | 1 | 0 |
| rptLastTenOpen | SELECT qryCaseListOpen.CaseID, qryCaseListOpen.CaseOpenDate,... | 5 | 0 |
| rpt_adj_address_label | SELECT [Personal Injury].Adjuster1, tblCase.CaseID, [Persona... | 1 | 0 |
| rptInvoiceComprehensiveAR | SELECT tblCase.CaseID, [Matter and AR].MatterID, [Matter and... | 3 | 0 |
| rpt_Compr_InvoiceStmtCur | SELECT [TB Time Keeping].[Bill Closed], [TB Time Keeping].IA... | 5 | 4 |
| rpt_Compr_InvoiceTKExCur | SELECT [TB Time Keeping].[Bill Closed], [TB Time Keeping].IA... | 5 | 4 |
| rptCriminalStatusNotesLog2 | tblNotes | 3 | 0 |
| rpt_MergeInvMatter | SELECT tblCase.CaseID, [Matter and AR].MatterID, [Matter and... | 3 | 0 |
| rpt_Comprehensive_Invoice2 | SELECT [TB Time Keeping].[Bill Closed], [TB Time Keeping].IA... | 5 | 4 |
| rpt_Trust_Chron_65 | qryTrustEntriesChron65 | 5 | 0 |
| rpt_TimeDetail_Comprehensive2 | qryInvoiceComprehensiveTimeDetail2 | 7 | 0 |
| rpt_TKLessTrust | qryInvoiceAttachRPT1 | 7 | 1 |
| rpt_Comprehensive_InvoiceTKEx2 | SELECT [TB Time Keeping].[Bill Closed], [TB Time Keeping].IA... | 5 | 4 |
| rpt_Comprehensive_InvoiceADV | SELECT [TB Time Keeping].[Bill Closed], [TB Time Keeping].IA... | 5 | 4 |
| rpt_OpenCases | qryCaseListOpen | 5 | 0 |
| rpt_Comprehensive_InvoiceTKEx2S | SELECT [TB Time Keeping].[Bill Closed], [TB Time Keeping].IA... | 5 | 4 |
| rpt_Comprehensive_InvoiceADVS | SELECT [TB Time Keeping].[Bill Closed], [TB Time Keeping].IA... | 5 | 4 |
| rpt_Comprehensive_InvoiceTKEx1 | SELECT [TB Time Keeping].[Bill Closed], [TB Time Keeping].IA... | 5 | 4 |
| rpt_Comprehensive_InvoiceStmt | SELECT [TB Time Keeping].[Bill Closed], [TB Time Keeping].IA... | 5 | 4 |
| rpt_Comprehensive_InvoiceTKEx | SELECT [TB Time Keeping].[Bill Closed], [TB Time Keeping].IA... | 5 | 4 |
| rpt_Open_Cases | qryTakeOff | 5 | 0 |
| rpt_Comprehensive_InvoiceTKEx1S | SELECT [TB Time Keeping].[Bill Closed], [TB Time Keeping].IA... | 5 | 4 |
| rpt_Comprehensive_InvoiceTKLessTrustCostAR | SELECT [TB Time Keeping].[Bill Closed], [TB Time Keeping].IA... | 5 | 4 |
| rpt_Comprehensive_InvoiceTKLessTrustRep | SELECT [TB Time Keeping].[Bill Closed], [TB Time Keeping].IA... | 5 | 4 |
| rpt_Comprehensive_InvoiceTKLessTrustRep2 | SELECT [TB Time Keeping].[Bill Closed], [TB Time Keeping].IA... | 5 | 4 |
| rpt_TimeDetail_Comprehensive | qryInvoiceComprehensiveTimeDetail | 7 | 0 |
| rpt_File_Folder_Label | qryFileFolderLabel | 1 | 0 |
| rpt_ftrustee_address_label | Bankruptcy | 1 | 0 |
| rpt_MergeInvTimeDetail | tblTimeTableDetail | 2 | 0 |
| Rpt_MergeInvTK | qryInvoiceRPT1 | 6 | 2 |
| rpt_trustee_address_label | Bankruptcy | 1 | 0 |
| rpt_opp_counsel_address_label | tblCase | 1 | 0 |
| rpt_TKExceedsTrust | qryInvoiceAttachRPT1 | 5 | 0 |
| rpt_Trust_Chron_35D | qryTrustEntriesChronRPT35D | 3 | 0 |
| rpt_Trust_Chron_65D | qryTrustEntriesChronRPT65D | 3 | 0 |
| rpt_Trust_Chron_95D | qryTrustEntriesChronRPT95D | 3 | 0 |
| rpt_Trust_Chron_35W | qryTrustEntriesChronRPT35W | 3 | 0 |
| rpt_Trust_Chron_65W | qryTrustEntriesChronRPT65W | 3 | 0 |
| rpt_Trust_Chron_95W | qryTrustEntriesChronRPT95W | 3 | 0 |
| rpt_Trust_Chron_95 | qryTrustEntriesChronRPT95 | 3 | 0 |
| rptBillingTotals | SELECT vwBillingTracker2.Tatty, Sum(vwBillingTracker2.Time_)... | 3 | 0 |
| rptClientNotes | SELECT tblCase.CaseID, tblCase.Last_Name, tblCase.First_Name... | 4 | 0 |
| rptComprehensiveTKStatement | qryInvoiceAttachComp | 6 | 0 |
| rptCriminalStatus | qryCrimStatus | 4 | 4 |
| rptCriminalStatusActionNeeded | SELECT TblActionNeeded.ActionNeededDet, TblActionNeeded.Acti... | 2 | 0 |
| rptCriminalStatusChargeNos | tbl_CtCaseNumbers | 2 | 0 |
| rptCriminalStatusUpcHrgs | SELECT tblHearingDate.CaseID, tblHearingDate.Hearing_Date, t... | 2 | 0 |
| rptInvoiceComprARCur | SELECT * FROM qry_InvoiceAR_curr;  | 3 | 0 |
| rptInvoiceComprehensiveTrust | qryInvoiceComprehensiveTrustCredit | 3 | 0 |
| rptInvoiceComprPymtsAR | SELECT tblCase.CaseID, [Matter and AR].MatterID, [Matter and... | 3 | 0 |
| rptInvoiceComprPymtsARCur | qry_InvoicePymts_curr | 3 | 0 |
| rptInvoiceComprTrustCur | qryInvoiceComprehensiveTrustCredit | 3 | 0 |
| rptLastWeekIntake | SELECT [TB Intakes].ID, [TB Intakes].[GI Last Name], [TB Int... | 5 | 0 |
| rptPersInjProviderBills | tblPersInjProv | 2 | 0 |
| rptPersInjStatusAction | SELECT TblActionNeeded.ActionNeededDet, TblActionNeeded.Acti... | 2 | 0 |
| rptPersInjStatusDemand | tblPersInjDemand | 2 | 0 |
| rptPersInjStatusLog | tblPersInjLog | 2 | 0 |
| rptPersInjuryStatus | qryPersInjStatus | 4 | 4 |
| rptPIStatusSOL | qryAttyTrustAcctsTOff | 5 | 0 |
| rptReceipt | qryReceipt | 5 | 0 |
| rptReceiptC | qryReceipt | 5 | 0 |
| rptReceiptR | tblReceipts | 5 | 0 |
| rptReceiptRec | tblReceipts | 5 | 0 |
| rptReconciliation | tblTakeOffMonth | 5 | 1 |
| rptTKReport | qryTimeKeeping | 5 | 0 |
| rptTKReport2 | qryTimeKeeping | 6 | 0 |
| Statement of Trust Account | qryStmtTrustRPT1 | 5 | 0 |

## Structured Form Inventory
**Extracted:** 94 forms

| Form | Record Source | Sections | Controls |
|------|---------------|----------|----------|
| frmClientLedger | SELECT vwfrmClientLedger.CaseID, vwfrmClientLedger.Last_Name... | 1 | 341 |
| frm_advanced_payments | qry_advanced_payments | 3 | 38 |
| frm_uncashed_trust_checks | qry_uncashed_trust_checks | 3 | 30 |
| frmActionNeededAll3 | qryActionNeededAll | 2 | 30 |
| frmBankruptcy | SELECT Bankruptcy.BankruptcyID, Bankruptcy.CaseID, Bankruptc... | 1 | 68 |
| frm_trust_summary | qry_trustStatements | 2 | 30 |
| frm_invoices_summary | SELECT vw_frm_invoices_summary.CaseID, vw_frm_invoices_summa... | 3 | 58 |
| frm_Billing_Tracker | qryBillingTracker | 2 | 19 |
| frmDispositions | qryDispos | 2 | 52 |
| frm_Billing_Tracker2 | qryBillingTracker2 | 2 | 36 |
| frmTimeTableDetailMerge | SELECT tblTimeTableDetail.Time_ID, tblTimeTableDetail.Tdate,... | 3 | 22 |
| frmChild | SELECT tblChild.Child_ID, tblChild.FamilyLaw_ID, tblChild.Ch... | 2 | 7 |
| frmActionNeeded | SELECT TblActionNeeded.ActionNeededID, TblActionNeeded.CaseI... | 1 | 9 |
| frmTakeOffReconciliation | qryTakeOff | 2 | 121 |
| frmAttyFeeGeneration | SELECT tblTakeOffMonth.*, tblTakeOffMonth.TakeOffDate FROM t... | 3 | 101 |
| frmActionNeededAll | qryActionNeededAll | 2 | 39 |
| frmActionNeededAll2 | qryActionNeededAll2 | 2 | 27 |
| frmCalls | tblCalls | 3 | 70 |
| frmLogin |  | 1 | 10 |
| frmClientReviews | SELECT tblCase.Last_Name, tblCase.First_Name, tblCase.Case_L... | 3 | 37 |
| frmAddUser | tblUsers | 1 | 11 |
| frmCalendarCheck | qryCalendarCheck | 2 | 39 |
| frmAdminLoginTK |  | 1 | 8 |
| frmPersInjDemand | tblPersInjDemand | 2 | 9 |
| frmHearingDate | SELECT tblHearingDate.HearingID, tblHearingDate.CaseID, tblH... | 2 | 21 |
| frmApplicationLoad |  | 0 | 0 |
| frmCaseList |  | 2 | 11 |
| frmAttyNotes | SELECT tblNotes.IDNotes, tblNotes.CaseID, tblNotes.NoteDate,... | 2 | 10 |
| frmBilling | SELECT Billing.ID, Billing.CaseID, Billing.[Balance Due Date... | 1 | 14 |
| frmBrowse |  | 1 | 5 |
| frmCaseListClosed | qryCaseListClosed | 2 | 34 |
| frmBrowse_BackEnd |  | 1 | 5 |
| frmReceipt | tblReceipts | 1 | 33 |
| frmClientsConflict | tblCase | 2 | 17 |
| frmUsers_Edit | SELECT * FROM tblUsers;  | 1 | 12 |
| frmCallsList | SELECT tblCalls.CFirstName, tblCalls.CLastName, tblCalls.CDa... | 3 | 39 |
| frmCaseListOpen subform | SELECT [qryCaseListOpen].[CaseID], [qryCaseListOpen].[CaseOp... | 1 | 30 |
| frmCaseListAll | qryCaseListAll | 2 | 38 |
| frmCaseListOpen | qryCaseListOpen | 2 | 39 |
| frmHomeAdmin |  | 3 | 12 |
| zfrmSelectCaseNum |  | 1 | 4 |
| frmConflictChk |  | 2 | 8 |
| frmPersInjuryStatusReport |  | 1 | 7 |
| frmOppPartyConflict | tblCase | 2 | 15 |
| frmFamilyLaw | SELECT [Family Law - Divorce].ID, [Family Law - Divorce].Cas... | 1 | 127 |
| frmOpenReport |  | 1 | 10 |
| frmCrimStatusReport |  | 1 | 11 |
| frmCtCaseNumbers | SELECT tbl_CtCaseNumbers.CtCaseNoID, tbl_CtCaseNumbers.CaseI... | 1 | 6 |
| frmDisposition | SELECT Disposition.DispoID, Disposition.CaseID, Disposition.... | 1 | 25 |
| frmHome |  | 3 | 38 |
| frmScanLocation | tblScans | 1 | 3 |
| frmHomeAdminLogin |  | 1 | 8 |
| frmOkAlert |  | 1 | 4 |
| frmIntakesConflicts | TB Intakes | 2 | 17 |
| frmSubProofOfClaims | SELECT ProofOfClaims.IDProofOfClaims, ProofOfClaims.IDBankru... | 2 | 23 |
| frmInvoiceSent | SELECT tbl_InvoiceSent.CaseID, tbl_InvoiceSent.InvSent, tbl_... | 2 | 18 |
| frmPersonalInjury2 | Personal Injury | 1 | 107 |
| frmMatter | SELECT vwMatterAndAR.MatterID, vwMatterAndAR.CaseID, vwMatte... | 1 | 17 |
| frmPersInjLog | SELECT tblPersInjLog.EventDate, tblPersInjLog.EventDescripti... | 2 | 10 |
| frmPersInjProvider | SELECT tblPersInjProv.PIProviderID, tblPersInjProv.ID, tblPe... | 2 | 21 |
| frmPersInjLog2 | tblPersInjLog | 2 | 6 |
| frmPersonalInjury | SELECT [Personal Injury].ID, [Personal Injury].CaseID, [Pers... | 1 | 109 |
| frmScansubform | SELECT tblScans.ScansID, tblScans.CaseID, tblScans.ScanLocat... | 1 | 2 |
| frmSourceAnalytics | qryCaseSourcesRPT1 | 3 | 45 |
| frmSubCH13Plans | SELECT CH13Plans.IDCH13Plans, CH13Plans.IDBankruptcy, CH13Pl... | 2 | 12 |
| frmSubPrevBankrupt | SELECT tblPrevBank.IDPrevBank, tblPrevBank.IDBankruptcy, tbl... | 2 | 10 |
| frmTakeOff | SELECT tblTakeOffMonth.*, tblTakeOffMonth.TakeOffDate FROM t... | 2 | 154 |
| frmTakeOff2 | SELECT tblTakeOffMonth.*, tblTakeOffMonth.TakeOffDate FROM t... | 2 | 9 |
| frmTakeOffSteps |  | 2 | 15 |
| frmTakeOffSubForm | SELECT vwfrmTakeOffSubForm.FileNumber, vwfrmTakeOffSubForm.N... | 2 | 84 |
| frmTakeOffSubForm_OLD | qryTakeOffStep2 | 2 | 77 |
| frmTakeOffSubForm2 | qryTakeOffStep2 | 2 | 63 |
| frmTakeOffSubForm3 | qryTakeOffStep2 | 2 | 75 |
| frmTakeOffTest | SELECT tblTakeOffMonth.*, tblTakeOffMonth.TakeOffDate FROM t... | 2 | 9 |
| frmTakeOffTotalFeesCosts | tblTakeOffMonth | 2 | 39 |
| frmTimeKeepingClosed | qryTimeKeepingClosed | 2 | 53 |
| frmTimeKeepingOpen | qryTimeKeepingOpen | 2 | 37 |
| frmTimeTableDetail | SELECT vwTimeTableDetail.Time_ID, vwTimeTableDetail.Tdate, v... | 3 | 31 |
| frmTKClose | qryTKClose1 | 2 | 71 |
| frmToBeClosed | qryToBeClosed | 2 | 32 |
| frmToBeScanned | qryToBeScanned | 2 | 31 |
| frmTrustAccount | SELECT vwTrustAccountTable.TrustAccountID, vwTrustAccountTab... | 1 | 17 |
| zfrmSelectCaseNum_Discount |  | 1 | 4 |
| frmTRUSTENTRIESCHRON | qryTrustEntriesChron | 3 | 68 |
| frmUpcoming Hearings | qryUpcomingHearings | 2 | 53 |
| frmUsers | SELECT * FROM tblUsers;  | 2 | 10 |
| frmYearWiseCaseList | SELECT TblCase.CaseID, TblCase.CaseOpenDate, [Last_Name] & "... | 2 | 41 |
| frmYesNoAlert |  | 1 | 5 |
| Intakes | TB Intakes | 3 | 59 |
| Time Keeping | qryTimeKeeping | 1 | 113 |
| zClient Ledger OLD | tblCase | 2 | 209 |
| zfrmFamilyLaw OLD | qryFamilyLaw | 2 | 164 |
| zfrmPersInjSOL | qrySOL | 2 | 27 |
| zfrmPersonalDetailsFamilyLaw | tblCase | 1 | 30 |

## JSON Layer Summary
| Artifact | Count |
|----------|-------|
| Structured forms (`extract/forms/*.json`) | 94 |
| Structured reports (`extract/reports/*.json`) | 99 |
| Query index (`extract/queries/index.json`) | 213 |
| VBA index (`extract/vba/index.json`) | 1200 procedures |
| Report lineage index (`extract/lineage/index.json`) | 99 reports |
| Linked tables | 88 |
| App manifest (`extract/app_manifest.json`) | 1 |

## Migration Hints
- **Startup object:** `frmClientLedger` (heuristic-first-form)
- **VBA red flags:**
  - `createObject` (23)
  - `eval` (1)
  - `fileSystem` (7)
  - `outputTo` (8)
  - `runSql` (46)
  - `setWarnings` (4)
  - `transferSpreadsheet` (1)
  - `transferText` (1)
- **High-risk areas:**
  - CreateObject calls found; external automation needs migration planning.
  - File-system access (FileSystemObject/Kill/MkDir) found; reconcile with deployment model.
  - 88 linked table(s) detected; back-end connectivity must be reproduced.

## Report Lineage Summary
**Generated:** 99 lineage report(s)

| Report | Trigger Paths | Queries | Tables | Confidence |
|--------|---------------|---------|--------|------------|
| rpt_TKTotalAdvance | 2 | 2 | 3 | high |
| rpt_Matter_Closing | 0 | 0 | 1 | medium |
| rptCriminalStatusNotesLog | 0 | 0 | 1 | medium |
| rpt_Trust_Closing | 0 | 1 | 1 | medium |
| rpt_Billing_Closing | 0 | 0 | 2 | medium |
| Accounts Receivable | 3 | 2 | 2 | high |
| rpt_Main_Closing | 6 | 0 | 1 | high |
| rpt_Disposition_Closing | 0 | 1 | 2 | medium |
| rpt_CaseNumber_Closing | 0 | 0 | 1 | medium |
| Case Sources and Revenue | 0 | 2 | 2 | medium |
| rpt_Comprehensive_InvoiceStmtS | 4 | 0 | 2 | high |
| rpt_Comprehensive_InvoiceTKEx3Costs | 8 | 0 | 2 | high |
| Copy Of Client Closing Sheet | 0 | 2 | 5 | medium |
| rpt_Compr_InvoiceADVCur | 2 | 0 | 2 | high |
| Client Closing Sheet | 0 | 2 | 5 | medium |
| rptInvoiceComprehensiveTrust2 | 0 | 1 | 2 | medium |
| Client_Trust_Accounts_for_PreTake_Off | 1 | 2 | 1 | high |
| Client_Trust_Accounts_for_Take_Off | 6 | 1 | 2 | high |
| rpt_Reconciliation sub | 0 | 1 | 1 | medium |
| rpt_Comprehensive_Invoice | 9 | 0 | 2 | high |
| New Invoice | 3 | 1 | 1 | high |
| Invoice | 49 | 1 | 1 | high |
| Invoice - No Balance Due | 5 | 1 | 1 | high |
| Invoice - Past Due | 9 | 1 | 1 | high |
| Invoice Attach - Hourly | 13 | 2 | 3 | high |
| rptPISOLList | 0 | 0 | 0 | medium |
| rpt_Comprehensive_InvoiceTKEx3CostsS | 4 | 0 | 2 | high |
| Invoice Attach - Hourly w Discount | 9 | 2 | 3 | high |
| rpt_Trust_Chron_35 | 3 | 1 | 2 | high |
| rptInvoiceComprehensiveAR2 | 0 | 0 | 2 | medium |
| Invoice2 | 5 | 1 | 1 | high |
| rpt_address_label | 6 | 0 | 1 | high |
| rpt_address_labelEx | 5 | 0 | 1 | high |
| rptLastTenOpen | 1 | 1 | 2 | high |
| rpt_adj_address_label | 1 | 0 | 2 | high |
| rptInvoiceComprehensiveAR | 0 | 0 | 2 | medium |
| rpt_Compr_InvoiceStmtCur | 2 | 0 | 2 | high |
| rpt_Compr_InvoiceTKExCur | 2 | 0 | 2 | high |
| rptCriminalStatusNotesLog2 | 0 | 0 | 1 | medium |
| rpt_MergeInvMatter | 0 | 0 | 2 | medium |
| rpt_Comprehensive_Invoice2 | 0 | 0 | 2 | medium |
| rpt_Trust_Chron_65 | 3 | 1 | 2 | high |
| rpt_TimeDetail_Comprehensive2 | 0 | 2 | 2 | medium |
| rpt_TKLessTrust | 2 | 2 | 3 | high |
| rpt_Comprehensive_InvoiceTKEx2 | 8 | 0 | 2 | high |
| rpt_Comprehensive_InvoiceADV | 9 | 0 | 2 | high |
| rpt_OpenCases | 1 | 1 | 1 | high |
| rpt_Comprehensive_InvoiceTKEx2S | 4 | 0 | 2 | high |
| rpt_Comprehensive_InvoiceADVS | 4 | 0 | 2 | high |
| rpt_Comprehensive_InvoiceTKEx1 | 8 | 0 | 2 | high |
| rpt_Comprehensive_InvoiceStmt | 9 | 0 | 2 | high |
| rpt_Comprehensive_InvoiceTKEx | 9 | 0 | 2 | high |
| rpt_Open_Cases | 0 | 2 | 1 | medium |
| rpt_Comprehensive_InvoiceTKEx1S | 4 | 0 | 2 | high |
| rpt_Comprehensive_InvoiceTKLessTrustCostAR | 0 | 0 | 2 | medium |
| rpt_Comprehensive_InvoiceTKLessTrustRep | 1 | 0 | 2 | high |
| rpt_Comprehensive_InvoiceTKLessTrustRep2 | 0 | 0 | 2 | medium |
| rpt_TimeDetail_Comprehensive | 0 | 2 | 3 | medium |
| rpt_File_Folder_Label | 1 | 1 | 2 | high |
| rpt_ftrustee_address_label | 1 | 0 | 1 | high |
| rpt_MergeInvTimeDetail | 0 | 0 | 1 | medium |
| Rpt_MergeInvTK | 0 | 1 | 1 | medium |
| rpt_trustee_address_label | 1 | 0 | 1 | high |
| rpt_opp_counsel_address_label | 1 | 0 | 1 | high |
| rpt_TKExceedsTrust | 2 | 2 | 3 | high |
| rpt_Trust_Chron_35D | 1 | 1 | 2 | high |
| rpt_Trust_Chron_65D | 1 | 1 | 2 | high |
| rpt_Trust_Chron_95D | 1 | 1 | 2 | high |
| rpt_Trust_Chron_35W | 1 | 1 | 2 | high |
| rpt_Trust_Chron_65W | 1 | 1 | 2 | high |
| rpt_Trust_Chron_95W | 1 | 1 | 2 | high |
| rpt_Trust_Chron_95 | 3 | 1 | 2 | high |
| rptBillingTotals | 1 | 0 | 1 | high |
| rptClientNotes | 2 | 0 | 2 | high |
| rptComprehensiveTKStatement | 1 | 1 | 2 | high |
| rptCriminalStatus | 3 | 1 | 2 | high |
| rptCriminalStatusActionNeeded | 0 | 0 | 1 | medium |
| rptCriminalStatusChargeNos | 0 | 0 | 1 | medium |
| rptCriminalStatusUpcHrgs | 0 | 0 | 1 | medium |
| rptInvoiceComprARCur | 2 | 2 | 1 | high |
| rptInvoiceComprehensiveTrust | 0 | 1 | 2 | medium |
| rptInvoiceComprPymtsAR | 2 | 0 | 2 | high |
| rptInvoiceComprPymtsARCur | 2 | 2 | 1 | high |
| rptInvoiceComprTrustCur | 2 | 1 | 2 | high |
| rptLastWeekIntake | 1 | 0 | 1 | high |
| rptPersInjProviderBills | 0 | 0 | 1 | medium |
| rptPersInjStatusAction | 0 | 0 | 1 | medium |
| rptPersInjStatusDemand | 0 | 0 | 1 | medium |
| rptPersInjStatusLog | 0 | 0 | 1 | medium |
| rptPersInjuryStatus | 4 | 1 | 2 | high |
| rptPIStatusSOL | 0 | 1 | 2 | medium |
| rptReceipt | 2 | 1 | 2 | high |
| rptReceiptC | 1 | 1 | 2 | high |
| rptReceiptR | 1 | 0 | 1 | high |
| rptReceiptRec | 0 | 0 | 1 | medium |
| rptReconciliation | 1 | 0 | 1 | high |
| rptTKReport | 2 | 1 | 2 | high |
| rptTKReport2 | 2 | 1 | 2 | high |
| Statement of Trust Account | 7 | 1 | 1 | high |

