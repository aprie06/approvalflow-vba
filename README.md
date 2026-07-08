# approvalflow-vba
Original VBA implementation of the ApprovalFlow timesheet approval system, anonymized for public sharing.
ApprovalFlow

**Automated intern timesheet approval system built in Outlook VBA and Excel**

ApprovalFlow is a production automation system I built for an internship program hosted at community colleges. The program manages 200 student interns placed at 85+ off-site employer organizations on a semi-monthly payroll schedule. Before this system existed, one HR staff member (me) was manually tracking down supervisor approvals across every pay period — a process that consumed 4 of every 5 working days.

I built this entirely on personal initiative. Automation is not part of my job description.

---

## Impact

| Metric |	Before |	After |

| Processing time per pay period |	~32 hours |	~12 hours |

| Time reduction |	—	| 62.5% |

| Hours saved annually |	—	| ~480 hours |

| Automated emails per pay period |	0	| 200–300 |

| Manual tracking errors |	Frequent |	Near zero |

---

## How It Works

Student interns submit timesheets in Banner/Student Information System (SIS). Banner sends a notification email to the shared payroll mailbox. ApprovalFlow intercepts that email, extracts the student's information from the HTML body, looks up the assigned supervisor from a roster in Excel, and forwards a formatted approval request to the supervisor — all within 60 seconds of submission.

Supervisors reply with `APPROVE`, `REJECT`, or `CORRECTIONS`. The system reads those replies, matches them back to the correct student record, logs the result, and updates a color-coded dashboard.

```

Banner notification email

&#x20;       |

&#x20;       v

[MoveItems] -- reads HTML body, looks up supervisor in Excel roster

&#x20;       |

&#x20;       v

Supervisor receives formatted approval email (Student CC'd for transparency)

&#x20;       |

&#x20;       v

Supervisor replies

&#x20;       |

&#x20;       v

[ProcessSupervisorReplies] -- logs reply, resolves Exchange DN to SMTP if internal

&#x20;       |

&#x20;       v

[UpdateSubmissionStatus] -- matches reply to Sent_Log, updates Submission_Tracking

&#x20;       |

&#x20;       v

[CreateEnhancedDashboard] -- color-coded status view for HR

&#x20;       |

&#x20;       v

HR approves verified timesheets in Student Information System (SIS) -> payroll processes automatically

```

---

Architecture

Files

File	Location	Purpose

`Student_Sup_email.xlsx`	Desktop	Master data file: roster, logs, tracking, dashboard

VBA module	Outlook (ThisOutlookSession + Module)	All automation logic

`Supervisor_Response_Log.xlsx`	Desktop	Secondary supervisor reply audit log

Sheets in Student_Sup_email.xlsx

Sheet	Purpose

`Updated_Supervisor email`	Student-supervisor roster. Cell J1 is the pay period anchor date.

`Sent_Log`	Record of every forwarded email: student name, email, supervisor, date sent

`Reply_Log`	Record of every supervisor reply parsed: response type, match method, timestamp

`Submission_Tracking`	One row per student per pay period: submission status, approval status

`Supervisor_Replies`	Raw reply data before processing

`Pay_Period_Summary`	Aggregate stats per pay period

Core Macros

Macro	Trigger	Purpose

`MoveItems`	New email in `Student_Time_Sheet` folder	Parses Banner notification, looks up supervisor, forwards approval email

`ProcessSupervisorReplies`	New email in shared mailbox Inbox	Logs incoming supervisor replies, moves them to `Supervisor_Replies` folder

`ReadSupervisorRepliesFolder`	Manual	Full rebuild of Reply_Log from `Supervisor_Replies` folder

`FastRebuildSentLog`	Manual	Rebuilds Sent_Log from emails in `Sent_super` folder

`LoadLookupDataIntoMemory`	Called by MoveItems	Loads roster into in-memory dictionary for fast lookups

`UpdateSubmissionStatus`	Manual  called by RefreshPayPeriodSummary	Matches Reply_Log entries to Sent_Log; writes APPROVED/REJECTED to Submission_Tracking

`PopulateSubmissionTracking`	Manual	Rebuilds Submission_Tracking from scratch (clear-then-repopulate — never append)

`CreateEnhancedDashboard`	Manual	Generates color-coded status view with duplicate flagging

`GeneratePayPeriodSummary`	Manual	Writes aggregate stats to Pay_Period_Summary

`RefreshPayPeriodSummary`	Manual	Runs UpdateSubmissionStatus then GeneratePayPeriodSummary sequentially

`DeduplicateSentLog`	Manual	Keeps newest entry per student; removes duplicates from resubmissions

`DeduplicateReplyLog`	Manual	Keeps newest reply per student

`RunSmartReminders`	Manual	Routes to the correct reminder sub based on current date relative to deadline

`SendApprovalWindowReminders`	Called by RunSmartReminders	Sends reminder to supervisors with 2+ days until deadline

`SendDeadlineDayReminders`	Called by RunSmartReminders	Sends urgent reminder on deadline day

`SendPastDeadlineReminders`	Called by RunSmartReminders	Sends final notice after deadline; includes payment delay warning

`StartNewPayPeriodComplete`	Manual — start of each pay period	Sets J1, clears logs, resets tracking for the new period

`ArchivePayPeriodWorkbook`	Manual — end of each pay period	Saves a dated copy of the workbook to archive folder

---

Pay Period Operating Procedure

Before the period opens

Close `Student_Sup_email.xlsx` completely

Run `StartNewPayPeriodComplete` — sets J1 to current pay period start date, clears prior period logs

Verify J1 on `Updated_Supervisor email` shows the correct pay period start date

During the period (daily)

Confirm Outlook is open and macros are enabled

Banner notifications arriving in `Student_Time_Sheet` are processed automatically by `MoveItems`

Supervisor replies arriving in the shared mailbox Inbox are processed automatically by `ProcessSupervisorReplies`

Run `RunSmartReminders` daily to send deadline-appropriate follow-ups

End of period

Run `DeduplicateSentLog`, then close the workbook

Run `DeduplicateReplyLog`, then close the workbook

Run `ReadSupervisorRepliesFolder` to rebuild Reply_Log from the folder, then close the workbook

Run `PopulateSubmissionTracking` — must clear and repopulate, never append

Run `UpdateSubmissionStatus`, then close the workbook

Run `CreateEnhancedDashboard`

Log into Student Information System (SIS) and manually approve all verified (GREEN) timesheets

Run `ArchivePayPeriodWorkbook`

>*Note:* Run each macro individually with the workbook closed between steps. The VBA opens its own invisible Excel instance — an already-open workbook causes stale data reads.

---

Known Issues and Fixes Applied

This section documents every significant bug encountered in production and how it was resolved. This is the primary reference for anyone inheriting or maintaining the system.

---

Silent failure: no emails forwarded, no log entries, no errors

Root cause: Cell J1 on `Updated_Supervisor email` holds a stale prior pay period date. All incoming Banner notifications fail the pay period date guard silently — no error is raised, nothing is logged.

Fix: Before running any macros, open the workbook and confirm J1 contains a date within the current pay period. `StartNewPayPeriodComplete` sets this automatically, but if J1 is wrong after running it, set it manually.

This is the most common failure mode in the system.

---

Office update broke external email delivery

Symptom: Forwarded emails appeared to send but were never delivered to external supervisor addresses after a Microsoft 365 update.

Root cause: The Office update changed how Outlook VBA resolves recipient addresses for external domains. `forwardItem.Send` was being called before recipients were resolved.

Fix: Added `forwardItem.Recipients.ResolveAll` immediately before `forwardItem.Send` in `MoveItems`.

---

Internal supervisors returned Exchange Distinguished Names instead of SMTP addresses

Symptom: Supervisor email matching failed for all `@district.edu` supervisors. Logged addresses showed Exchange DN strings (`/O=EXCHANGELABS/...`) instead of real email addresses.

Root cause: `oMail.SenderEmailAddress` returns the Exchange Distinguished Name for internal Exchange users, not their SMTP address.

Fix: Added `GetSMTPAddress` helper function using `oMail.Sender.GetExchangeUser.PrimarySmtpAddress` to resolve Exchange users to their real SMTP address. Falls back to `SenderEmailAddress` for external senders.

---

Supervisor replies mismatched when supervisor manages multiple students

Symptom: All replies from a supervisor who oversees multiple students were matched to the same student, leaving the others unmatched.

Root cause: The matching algorithm weighted reply time proximity too heavily. The first student in the Sent_Log for that supervisor always scored highest.

Fix: Reweighted the scoring logic: name match score is weighted at 2, time proximity score is capped at 1. Name match now dominates. CC-based matching (extracting the student's `@district.edu` address from the reply CC field and matching to Sent_Log column 6) was added as the primary strategy; supervisor email plus name scoring became the fallback.

---

isReply gate silently dropped valid approval emails

Symptom: Some supervisor replies were never logged. No error, no log entry.

Root cause: `isReply` required the subject line to contain keywords ("Timesheet", "Internship Program Name", "Approval") in addition to body keywords. Replies forwarded through certain email clients had modified subjects that stripped these keywords.

Fix: Removed subject keyword requirement from `isReply`. Detection now relies on body keywords (APPROVE, REJECT, CORRECTIONS) only.

---

DetermineResponseType misclassified approvals as REJECTED

Symptom: Supervisor replies containing "APPROVE" were logged as REJECTED.

Root cause: The forwarded approval request email (quoted in the reply body) contained the word "REJECTED" in the instructions section ("If the hours are incorrect, reply with REJECTED"). The function was scanning the full email body including quoted content.

Fix: Strip quoted content from the reply body before keyword matching. Quoted content is identified by separators: "From:", a line of underscores, "-----Original Message-----", "Sent:", or "On [date]".

---

PopulateSubmissionTracking silent append bug caused 2x row overcounting

Symptom: `Submission_Tracking` showed 368 rows when the correct student count was 181.

Root cause: `PopulateSubmissionTracking` prompted the user with a yes/no dialog before archiving. The `vbNo` path had no explicit handler, causing the function to fall through and append new rows to existing ones rather than replacing them.

Fix: `PopulateSubmissionTracking` now always clears `Submission_Tracking` before repopulating, regardless of the archive prompt response. The `vbNo` path explicitly clears and repopulates without archiving. The `vbCancel` path exits without making any changes.

>*This is the system's most dangerous silent data corruption risk. PopulateSubmissionTracking must always clear before repopulating — never append.*

---

FastRebuildSentLog dropped valid timesheets submitted on deadline day

Symptom: FastRebuildSentLog returned 115 records when the correct count was 179. Timesheets submitted on the last day of the pay period were missing.

Root cause: A date filter using `ExtractTimesheetStatusDate` was filtering out emails where the Banner status date matched the submission date rather than a date within the pay period. Banner stamps the submission date, not the pay period date — a timesheet submitted on May 5 for the April 16-30 period had a status date of May 5, which fell outside the pay period window.

Fix: Removed the Banner status date filter entirely. All emails in `Sent_super` are now processed without date filtering.

---

Date mismatch: prior period approvals matched to current period submissions

Symptom: Dashboard showed entries like "Submitted 4/15, Approved 4/1" — logically impossible.

Root cause: `LogSupervisorReply` had no date guard. Supervisor replies from the prior pay period were being matched against current period submissions.

Fix: Added a date guard inside the `For sentRow` loop in `LogSupervisorReply` that skips any reply email with a received date earlier than `GetPayPeriodStartDate()`. `GetPayPeriodStartDate` reads from J1 on `Updated_Supervisor email`.

---

Outlook freeze during FastRebuildSentLog

Symptom: Outlook froze and became unresponsive when FastRebuildSentLog processed more than a few dozen emails.

Root cause: MSHTML DOM parsing (`htmlDoc.Body.innerHTML`) inside a per-email loop caused Outlook to freeze. Each DOM parse instantiated a new COM object while the previous one had not been released.

Fix: Replaced MSHTML DOM parsing with a plain string `ExtractEmailFromHTMLBody` function that counts `<td>` elements using string operations. A single batch array write (`Resize(resultCount, 8).Value = results`) replaced cell-by-cell writes inside the loop.

---
Hardcoded username paths will break under a different Windows account

Symptom (future risk): `ArchivePayPeriodWorkbook` and `ViewArchivedPayPeriods` contain hardcoded paths using `ApprovalFlow_Admin` as the username. These will fail when the system is handed off to a successor with a different Windows account.

Fix: Replace all hardcoded `C:\Users\ApprovalFlow_Admin\` paths with `Environ("USERPROFILE") & "\"`.

Note: this fix has been applied in the anonymized version of the code published in this repository.
---

Critical Operating Rules

These rules exist because each one was learned from a production incident.

J1 must be set before running any macros. If J1 on `Updated_Supervisor email` holds a stale date, the system silently rejects all incoming emails with no error or log entry.

Student_Sup_email.xlsx must be fully closed before running any macro. The VBA opens its own invisible Excel instance. An already-open workbook causes conflicts and stale data reads.

PopulateSubmissionTracking must always clear before repopulating. Never allow it to append to existing rows. The append bug is the system's most dangerous silent data corruption risk.

Run macros individually with the workbook closed between steps. Sequential macro runs (e.g., via RefreshPayPeriodSummary) can produce stale data reads. Run each sub individually as the workaround.

Use Display instead of Send when testing. Always verify email content with `.Display` before switching to `.Send` for live runs.



## System Context and Authorship

* **Program:** Student internship program, Multi-Campus Community College
* **Developer:** aprie06
* **Development period:** January 2024 to present
* **Languages/tools:** Outlook VBA, Excel VBA, MSHTML, Banner (Ellucian) via HRIS SQL queries
* **All subsequent development, logic, debugging, and enhancements:** aprie06





