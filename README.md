# Mini LIMS - An Excel Prototype for QC Sample and Result Management

A Mini LIMS prototype in Excel - 8 sheets covering sample intake, test results, and specifications, with cross-sheet formula references, automated OOS classification, and a dashboard summarising status and OOS rates.

## Purpose

Pharmaceutical QC labs rely on LIMS to maintain traceability of incoming raw materials, in-process samples, and finished products across the full testing lifecycle. Controlled workflows and audit trails reduce manual error and preserve data integrity.
This prototype reproduces core LIMS concepts in Excel across 8 linked sheets - sample log, controlled test methods, product-specific specifications, automated results classification, chain of custody, and audit trail. A dashboard summarises current operational state - total sample count, status distribution, OOS rate, and OOS breakdown by product - giving a QC Manager a single-screen view of lab status.
This workbook is a learning and portfolio artefact, not production software. It does not implement the automated event capture or user authentication required in a production LIMS.

## Workbook Structure

| Sheet | Purpose |
| --- | --- |
| Samples | Master log of all samples - one row per sample with intake details, logging analyst, storage requirements and sample status. |
| Test Methods | Reference catalogue of available test procedures, with SOP versions, equipment, and units. |
| Specifications | Acceptance criteria for each product–test method combination, with lower and upper limits, units, and effective dates. |
| Results | Test results event log - one row per test performed, with auto-populated specifications, formula-driven Pass/OOS classification, and analyst attribution. |
| Chain of Custody | Timestamped record of every physical handoff - sample movement between analysts, locations, and storage states. |
| Audit Trail | Record of every data change across sheets - with Old/New values, reason for change, and approver |
| Analysts | Reference list of QC staff - initials, full names, roles, and training records - used as the source for analyst dropdowns across the workbook. | 
| Dashboard | Summary dashboard showing total sample count, status distribution, OOS rate, and OOS breakdown by product. |

## Key Design Decisions

### Decision 1: Specifications structured around product–test method combinations, not per sample

Specifications are a property of the product and its test method, not of any individual sample. If specs were per-sample, the duplicated spec would create inconsistency risk and make it harder to maintain a single approved specification sheet. In other words, there wouldn't be a single source of truth for the same product–test pairing since the values would be scattered across rows. Each sample having its own spec would also make the workbook harder to audit and maintain consistently.
*Tradeoff: this assumes specs change rarely. When they do, the workbook uses a Status field (Active/Superseded) to retain spec history rather than deleting old specs.*

### Decision 2: Use of multi-key XLOOKUP for spec retrieval

Specification retrieval depends on both the product name and the test method. I used an XLOOKUP formula, which would call the product and its respective test method.  The pass or fail of a sample in the results only works based on these acceptance limits, which in turn depend on the correct spec, which in turn needs the correct product name and test method. The formula matches the combined string (product, test method) against the combined columns and returns the matching units, lower limit and upper limit values based on this string match.  *Tradeoff: For the formula to work, exact word matching is essential. In a proper LIMS software, the data are predefined and already enforced at the database level.* 

### Decision 3: Layered defensive guards on the Pass/Fail formula

One of the key things I noticed while completing and proofreading the workbook was that when non-numeric characters were used in the measured value section, it still gave an OOS output, which is wrong technically, and this would affect the OOS rate in the dashboard sheet. Initially, the formula was set up for 3 options: empty value returns blank, no matching spec returns blank and numeric values are compared against the limits, and returns pass/ OOS. Then, after the error was found, the NOT(ISNUMBER()) check was added, and the final formula checks first whether the cell is empty, then checks if it's a number, then checks if there's a spec, and only then gives the Pass/OOS output. 
*Tradeoff: The formula catches invalid input downstream rather than preventing it at entry. In a production LIMS, numeric fields reject non-numeric input at the cell level — eliminating the need for defensive checks. Data Validation on the Measured Value column could improve this in Excel.*

### Decision 4:  Named ranges for cross-sheet dropdowns

Across multiple sheets, the same lists of values - analysts, sample IDs, methods- needed to appear in dropdowns. Copying each list into every data validation box would create maintenance and integrity issues, so I used named ranges (`AnalystList`, `SampleIDList`, `MethodIDList`, etc.)  pointing back to master tables instead, ensuring a controlled vocabulary and a single source of truth. Since the dropdowns accept values only from their master list, spelling inconsistencies across sheets can't happen.
*Tradeoff: Each time I had to add a new analyst or a new sample, I had to extend the named range manually. In a production LIMS, master lists are database tables that expand automatically - there's no fixed range to maintain, so dropdowns stay in sync with the underlying data without manual intervention.*

### Decision 5: Simulating an Audit Trail

An audit trail is a record of changes made to data in the system - sample status updates, result corrections, spec revisions, and so on. In production LIMS, it automatically captures who is making the change, what changes have been made, the time of making the change, the reason behind the change and who approved the change. The audit trail sheet in my workbook demonstrates a sample trail but has its own limitations. The changes are not auto-captured but manually populated to mirror what a real LIMS produces. Automated capturing would require building an external automation layer, which is beyond the scope of this prototype.

## Limitations and Assumptions

- **Scalability** - This workbook was designed as a small-scale prototype and is not intended to support the larger data volumes, multiple users, or workflow complexity of a production LIMS.
- **Audit Trail** - Audit trail entries are manually populated for demonstration purposes rather than automatically captured by the system as they would be in a production LIMS.
- **User Authentication** - The workbook does not include user authentication, electronic signatures, or role-based access controls that would typically be present in a production LIMS.
- **Specification Data** - The sample specifications are included for illustrative purposes only, with some Ibuprofen values referenced from publicly available USP information, while Paracetamol and Ozempic values were not formally verified for regulated use.

## How to Use

Open the workbook from the Dashboard sheet to view the summary KPIs, charts, and recent audit trail activity. The remaining tabs represent different parts of the QC workflow, including sample registration, test methods, specifications, result entry, and chain of custody tracking. Charts and KPI values update automatically based on the underlying workbook data and formulas. The workbook is best viewed in the desktop version of Microsoft Excel.

