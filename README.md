# Restaurant Billing & Feedback Automation

A trigger-driven restaurant billing, feedback-collection, and customer-recovery system built entirely with **Google Sheets, Google Apps Script, Google Forms, and Gmail**.

The system turns a single Google Sheet into an automated workflow for:

- Sending invoice emails
- Collecting customer feedback through a prefilled Google Form
- Sending automated feedback reminders
- Detecting low customer-satisfaction scores
- Alerting the restaurant owner when intervention is required
- Detecting customer email replies
- Tracking the complete lifecycle of each invoice

> **Built end-to-end with Claude (Anthropic) from a structured engineering specification written by me. I own the specification, testing, debugging, and review process.**



---

## 1. Problem

Small restaurants often handle invoicing and customer feedback manually.

This creates several operational problems:

- Invoices may be sent late or inconsistently.
- Customers are not always asked for feedback.
- Feedback reminders require manual follow-up.
- Unhappy customers can go unnoticed.
- Customer complaints are difficult to track systematically.
- There is often no simple record of which invoice requires action.

The goal of this project was to build a **low-cost, spreadsheet-based automation system** that a small restaurant could operate without needing a dedicated CRM or external automation platform.

---

## 2. Solution

The restaurant owner only needs to enter the invoice information and change the row's **Status** to `Ready`.

From there, the system manages the workflow automatically.

```text
                         ┌──────────────────┐
                         │   Status = Ready │
                         └────────┬─────────┘
                                  │
                                  ▼
                         ┌──────────────────┐
                         │  Invoice Email   │
                         └────────┬─────────┘
                                  │
                                  ▼
                    ┌──────────────────────────┐
                    │ Feedback Email + Prefill │
                    │       Form Link          │
                    └────────────┬─────────────┘
                                 │
                  ┌──────────────┴──────────────┐
                  │                             │
                  ▼                             ▼
          Customer submits              No response
             feedback                  after interval
                  │                             │
                  ▼                             ▼
       Calculate mean of 5              Reminder 1
             ratings                         │
                  │                          ▼
          ┌───────┴───────┐             Reminder 2
          │               │                  │
      Score ≥ 3.0      Score < 3.0          ▼
          │               │             Exhausted
          ▼               ▼
     Thank-you       Owner alert +
       email         follow-up flag
```

Customer email replies are separately detected and flagged for the restaurant owner.

---

# 3. Core Workflow

### Step 1 — Create the invoice

The restaurant owner enters the invoice and customer information into the Google Sheet.

### Step 2 — Set status to `Ready`

Changing the status triggers the invoice workflow.

### Step 3 — Invoice email

The system generates and sends an HTML invoice email.

### Step 4 — Feedback request

The customer receives a feedback email containing a **prefilled Google Form link**.

The invoice number is prefilled so the customer does not need to manually enter it.

### Step 5 — Automated reminders

If the customer does not respond, the system sends:

1. Reminder 1
2. Reminder 2
3. Marks the feedback process as `Exhausted`

### Step 6 — Feedback processing

The system calculates the mean score from the five rating questions.

```text
Score >= 3.0
        ↓
Thank-you email

Score < 3.0
        ↓
Customer-recovery email
        +
Owner alert
        +
Need to Reply = Yes
```

### Step 7 — Customer replies

The system periodically scans the sending Gmail account for replies and flags relevant invoices for owner attention.

---

# 4. Architecture

| Layer | Technology | Purpose |
|---|---|---|
| Interface | Google Sheets | Invoice entry and workflow control |
| Validation | Sheets Data Validation | Prevent invalid status/state values |
| Feedback | Google Forms | Customer feedback collection |
| Delivery | `MailApp` | Sending automated emails |
| Reply Detection | `GmailApp` | Detecting customer replies |
| Orchestration | Apps Script triggers | Running automation automatically |
| Concurrency | `LockService` | Preventing overlapping executions |
| Presentation | HTML email | Customer-facing invoice and feedback emails |
| Storage | Google Sheets | System of record |

### Trigger architecture

The system uses:

- `onEdit` — reacts to `Ready` / `Stop`
- Form-submit trigger — processes feedback
- 1-minute sweep — reconciles missed events and handles scheduled work
- 5-minute reply scan — detects customer email replies
- Daily 9 AM trigger — sends owner summary
- `onOpen` — adds the custom spreadsheet menu

The exact trigger installation is handled by the setup process.

---

# 5. State Machine

The automation is designed around explicit states rather than relying only on timestamps.

A simplified lifecycle is:

```text
Draft
  │
  ▼
Ready
  │
  ▼
Invoice Sent
  │
  ▼
Feedback Requested
  │
  ├───────────────┐
  │               │
  ▼               ▼
Response       No Response
  │               │
  │          Reminder 1
  │               │
  │               ▼
  │          Reminder 2
  │               │
  │               ▼
  │           Exhausted
  │
  ├── Score >= 3.0 → Thank You
  │
  └── Score < 3.0  → Recovery + Owner Alert
```

A `Stop` state is also supported and prevents further automated customer communication.

---

# 6. Important Engineering Decisions

The project was designed around reliability rather than simply making the happy path work.

| Engineering decision | Purpose |
|---|---|
| **Claim-before-send** | Records the next state before sending a follow-up, reducing duplicate sends during retries |
| **Immediate `Stop` re-check** | Prevents an already-stopped invoice from sending another message |
| **Missed-response reconciliation** | Allows the periodic sweep to recover from a missed form-submit event |
| **Idempotent setup** | Re-running setup does not intentionally create duplicate infrastructure |
| **Per-invoice tracking** | Allows the same customer to have multiple independent orders |
| **Quota guard** | Prevents the system from silently exceeding the available email quota |
| **HTML escaping** | Prevents customer-entered text from being interpreted as HTML |
| **Placeholder guard** | Prevents placeholder configuration values from reaching customers |
| **Separate scheduling add-on** | Allows scheduling functionality to be added without modifying the tested core |
| **Execution time guards** | Large operations can defer remaining work to a later sweep instead of exceeding Apps Script's execution limit |

---

# 7. Repository Structure

```text
restaurant-billing-feedback-automation/
│
├── Code.gs
│   └── Core automation
│       ├── Configuration
│       ├── Sheet setup
│       ├── Form creation
│       ├── Email generation
│       ├── Workflow logic
│       ├── Validation
│       └── Trigger installation
│
├── FollowUpSchedule.gs
│   └── Add-on functionality
│       └── Sheet-controlled follow-up scheduling
│
├── docs/
│   ├── Setup_Guide.md
│   └── screenshots/
│
└── README.md
```

---

# 8. Quick Start

## Prerequisites

You need:

- A Google account
- Google Sheets
- Google Forms
- Gmail
- Google Apps Script

No external API, paid automation platform, or third-party database is required.

---

## Installation

### 1. Create the spreadsheet

Create a new Google Sheet.

Open:

```text
Extensions → Apps Script
```

---

### 2. Add the Apps Script files

Add:

```text
Code.gs
FollowUpSchedule.gs
```

Copy the corresponding source code into each file.

---

### 3. Configure the system

Update the `CONFIG` section with the restaurant's details.

Typical configuration includes:

```text
Restaurant name
Restaurant address
Phone number
Owner email
Sender configuration
Escalation / representative message
```

Do not leave required configuration values as placeholders.

---

### 4. Run the setup

Reload the spreadsheet.

A custom menu should appear:

```text
Billing Automation
```

Select:

```text
Billing Automation
    → Run Full Setup
```

Authorize the required Google permissions when prompted.

If required by the implementation, run the setup process again after authorization.

---

### 5. Install scheduled follow-ups

Run:

```text
installScheduledFollowUps
```

The scheduling add-on provides sheet-controlled follow-up timing.

---

### 6. Configure the Google Form

The automation creates/configures the feedback form.

Some Google Forms presentation features may require manual configuration because they are not fully exposed through the Apps Script Forms API.

Follow the instructions in:

```text
docs/Setup_Guide.md
```

---

### 7. Test the workflow

Create a test invoice using your own email address.

Change:

```text
Status → Ready
```

The automation should begin the invoice → feedback → reminder workflow.

---

# 9. Testing

The system includes an **8-case manual test plan**.

| Test | Scenario | Expected result |
|---|---|---|
| 1 | Idempotent setup | Re-running setup does not intentionally duplicate infrastructure |
| 2 | Happy path | Invoice and feedback workflow executes correctly |
| 3 | HTML escaping | Customer-entered text is safely rendered |
| 4 | Validation failures | Invalid configuration/data is rejected |
| 5 | Reminders | Reminder stages execute and eventually reach `Exhausted` |
| 6 | High score | Customer receives thank-you email |
| 7 | Low score | Customer receives recovery message and owner receives alert |
| 8 | Stop + reply detection | Stopped invoices stop further automation and replies are flagged |



---

# 10. Reliability & Safety Measures

### 10.1 Claim-before-send

For automated follow-ups, the system first claims the stage in the sheet and then performs the email operation.

Conceptually:

```text
Check stage
    ↓
Claim stage
    ↓
Save state
    ↓
Send email
```

This reduces the possibility of the same follow-up being sent twice if a trigger retries.

---

### 10.2 Stop protection

The system checks for `Stop` immediately before sending automated communication.

This is important because a row may change between:

```text
Trigger starts
        ↓
System processes row
        ↓
User changes status to Stop
```

The final check reduces the chance of sending after the user has explicitly stopped the workflow.

---

### 10.3 Missed-event reconciliation

The system does not rely exclusively on event triggers.

A periodic sweep checks the sheet and reconciles work that may have been missed by an event trigger.

This provides a second layer of reliability:

```text
Event trigger
     +
Periodic reconciliation
     ↓
More resilient workflow
```

---

### 10.4 Concurrency control

`LockService` is used to prevent overlapping executions from processing the same workflow simultaneously.

Conceptually:

```text
Trigger A ──┐
            ├── Lock ── Process
Trigger B ──┘
               │
               ▼
             Release
```

---

### 10.5 HTML escaping

Customer-controlled values are escaped before being inserted into HTML email templates.

This prevents values such as names or other sheet-entered fields from being interpreted as arbitrary HTML.

---

# 11. Email Workflow

The system generates HTML emails for the customer-facing communication.

The primary communication stages are:

```text
Invoice
   ↓
Feedback Request
   ↓
Reminder 1
   ↓
Reminder 2
```

Depending on the customer's response:

```text
Positive / neutral response
        ↓
Thank-you email
```

or:

```text
Low score
   ↓
Customer-recovery message
   ↓
Owner alert
   ↓
Need to Reply = Yes
```

The low-score workflow uses the configured representative message rather than automatically attempting to resolve the customer's complaint.

---

# 12. Feedback Scoring

The feedback form contains five rating questions.

The overall score is calculated as the arithmetic mean:

```text
Score =
(Rating 1 + Rating 2 + Rating 3 + Rating 4 + Rating 5) / 5
```

The current decision threshold is:

```text
Score >= 3.0
    → Thank-you path

Score < 3.0
    → Recovery path
```

The threshold is an implementation rule, not a claim about what constitutes customer satisfaction universally.

---

# 13. Duplicate & Repeat Customer Handling

The system tracks the workflow **per invoice rather than per customer**.

This is important because the same customer may visit multiple times.

For example:

```text
Customer: Rahul

INV-1001 → Feedback workflow
INV-1002 → Separate feedback workflow
INV-1003 → Separate feedback workflow
```

The Google Form is therefore not configured around a one-response-per-account model.

Invoice-level tracking provides the identifier required to distinguish separate transactions.

---

# 14. Quota Handling

The system includes an email-quota guard.

For free Gmail accounts, the practical daily sending limit is approximately **100 recipients/day**, while Google Workspace accounts generally have higher limits.

Exact quotas depend on Google's current account and service limits.

The automation is designed to **refuse/log work when the available quota is insufficient rather than silently dropping emails**.

For production deployments, administrators should verify the current limits applicable to their Google account.

---

# 15. Known Limitations

This is a prototype designed around Google's native services.

### Gmail limits

Email sending is subject to Google's account-specific quotas.

### Reply detection

Reply detection depends on the customer response reaching the Gmail account used by the automation.

Automated replies may also appear as replies and therefore require appropriate handling.

### Email rendering

HTML email rendering varies between email clients.

The current implementation has been manually checked against selected Gmail environments, but client compatibility cannot be guaranteed universally.

### Single location

The current architecture assumes:

```text
1 restaurant
1 spreadsheet
1 primary sender
```

### Single system of record

The Google Sheet is the central operational database.

There is currently no external relational database or dedicated CRM backend.

### Google platform dependency

The automation depends on Google Apps Script, Sheets, Forms, and Gmail behaviour and quotas.

Changes to Google's platform APIs or account policies may require maintenance.

---

# 16. Roadmap

Potential future improvements include:

### Configuration

- Settings sheet instead of editing `CONFIG`
- Admin configuration interface
- Better validation of configuration values

### Billing

- Item-level invoice lines
- Quantity and unit-price calculations
- Tax support
- Discounts
- Automatic invoice numbering
- PDF invoice generation

### Communication

- WhatsApp integration
- Additional reminder channels
- More flexible communication templates

### Analytics

- Feedback dashboard
- Customer satisfaction trends
- Response-rate analysis
- Recovery-rate analysis
- Power BI / Looker Studio integration

### Scalability

- Multi-outlet support
- Outlet-specific configuration
- Centralized reporting
- Database-backed storage

---

# 17. AI-Assisted Development

This project was generated with **Claude (Anthropic)**.

However, the project was not created by giving an AI a vague instruction such as:

> "Build a restaurant email automation."

The development process started with a structured engineering specification.

My role included:

- Defining the business requirements
- Designing the workflow
- Defining the data/state model
- Specifying constraints
- Defining acceptance criteria
- Testing the generated implementation
- Reporting observed failures
- Reviewing proposed fixes
- Verifying behaviour against the specification

The AI generated implementation code, but generated code was treated as **unverified until tested**.

---

# 18. Specification Approach

The specification included several constraints designed to reduce ambiguity.

| Technique | Application |
|---|---|
| **Role definition** | Specified Apps Script engineering and email-template requirements |
| **Anti-assumption rule** | Unspecified business values were not to be invented |
| **Fixed output structure** | Required code, documentation, checklist and test plan |
| **Exact schema** | Defined sheet columns, validation values and formatting |
| **State machine** | Defined workflow states and transitions |
| **Platform constraints** | Required unsupported Google Forms features to be documented instead of fabricated |
| **Non-functional requirements** | Included locking, idempotency, quota handling and escaping |
| **Acceptance criteria** | Converted requirements into testable behaviours |
| **Design requirements** | Specified the intended restaurant-facing visual style |

This approach made the AI behave more like an implementation assistant working from an engineering brief rather than an unrestricted code generator.

---

# 19. Iterative Debugging

The implementation was tested through multiple iterations.

The important distinction was that failures were reported based on **observed behaviour**, rather than assuming what the code was doing internally.

| Observed behaviour | Investigation / instruction | Result |
|---|---|---|
| Setup exceeded execution time | Execution log was provided for investigation | Blocking interaction was identified and replaced with a non-blocking approach |
| Placeholder sender configuration blocked sending | Configuration semantics were clarified | Sender fallback behaviour was implemented |
| Repeat customers could not complete multiple workflows | Business requirement was clarified | Form behaviour and invoice-level tracking were adjusted |
| Follow-ups needed to respect a sheet-defined time | Scheduling requirement was specified | Scheduling was separated into an add-on file |
| Low-score customers required human follow-up | Exact business message was specified | Recovery path was updated without changing the positive path |

The important development loop was:

```text
Specification
      ↓
AI-generated implementation
      ↓
Run
      ↓
Observe actual behaviour
      ↓
Record failure
      ↓
Describe requirement / failure
      ↓
Modify implementation
      ↓
Retest
```

---

# 20. What I Verify Myself

AI-generated code is not treated as proof that the system works.

Before considering the prototype complete, I verify:

- The complete 8-case test plan
- Execution logs
- Trigger behaviour
- Duplicate-send protection
- Stop behaviour
- Feedback processing
- Low-score escalation
- Reply detection
- Gmail rendering
- Mobile rendering
- Dark-mode behaviour
- Configuration validation
- README claims against actual implementation

The README should therefore be updated whenever implementation behaviour changes.

---

# 21. Honest Disclosure

This repository contains **AI-generated code**.

Claude was used as the implementation assistant.

The human contribution was the engineering specification, architecture, requirements, testing, debugging, review, and acceptance of the implementation.

The code should be reviewed and tested before being used with real customer data.

Do not treat AI-generated code as automatically secure, correct, or production-ready.

---


## Project Summary

This project demonstrates more than the ability to generate Apps Script code.

It demonstrates a workflow for turning a business requirement into an automated system:

```text
Business Problem
       ↓
Requirements
       ↓
System Design
       ↓
State Model
       ↓
Implementation
       ↓
Testing
       ↓
Failure Analysis
       ↓
Iteration
       ↓
Documented System
```

The objective is not simply to show that AI can write code.

It is to demonstrate that AI can be directed through a **structured engineering process**, while the resulting system is still tested and reviewed by a human.
