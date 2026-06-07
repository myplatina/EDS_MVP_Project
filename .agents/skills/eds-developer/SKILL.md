\# Antigravity Skill: EDS MVP Code Refactoring Standard



\## Trigger Conditions

\* Activated when modifying Google Apps Script (`.js`/`.gs`), React Frontend components, Vercel deployment configs, or updating project versions.



\## Architectural Constraints

1\. \*\*Ledger Integrity:\*\* Never generate functions that wipe out or directly mutate the '2. 종목현황' or '6. 배당내역' sheets unless synchronizing verified external prices (`refreshKrxPricesToMainSheet`).

2\. \*\*GAS Timeout Mitigation:\*\* When modifying `refreshAllKrxDailyChartsFromKis()`, you must implement chunking or pagination to prevent the 6-minute GAS timeout execution limit.

3\. \*\*PWA First:\*\* All React UI components must support mobile-first responsive design, matching the MMORPG balance lead's preference for dense, precise, and high-contrast data presentation.



\## Automated Verification Protocol (Sandbox)

\* After refactoring code, run the local build/test compiler inside the Antigravity secure sandbox.

\* Use the virtual browser to verify that the 'Rebalancing' and 'Risk Monitoring' cards render without component breaking or layout shifting.

