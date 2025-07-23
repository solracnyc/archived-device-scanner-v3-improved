# Performance Analysis & Optimization Research

## Current Performance Profile (v3.2.1)

### Benchmark Results
- **Large Scale Test**: 5,000 accounts scanned successfully
- **Total Duration**: Several hours with automatic pause/resume cycles
- **Execution Pattern**: 4.5-minute cycles with 60-second continuation delays
- **Processing Method**: Sequential (one account at a time)
- **API Usage**: One `AdminDirectory.Mobiledevices.list()` call per email address

### Current Architecture

```javascript
// Current approach - Sequential processing
for each email in accountList:
  getUserDevices(email)  // Individual API call
    ↓
  AdminDirectory.Mobiledevices.list('my_customer', {
    query: 'email:' + email,
    maxResults: 100,
    projection: 'FULL'
  })
```

### Identified Bottlenecks

1. **Sequential Processing**: Accounts processed one by one
2. **Individual API Calls**: One REST call per email (5,000 calls for 5,000 accounts)
3. **Full Data Retrieval**: `projection: 'FULL'` when preview only needs basic info
4. **Conservative Timing**: Fixed 60-second continuation delays
5. **No Caching**: Re-scans all accounts even if recently checked

## Optimization Research Prompt

Use this comprehensive research prompt with another AI to investigate performance improvements:

```markdown
You are a senior Google Workspace solutions architect and performance-engineering specialist.  
Your task is to design a technical research plan—and deliver concrete, evidence-based recommendations—for dramatically accelerating the "Minimal Archived Device Scanner" Google Apps Script project (v3.2.0).

─────────────────────────────
📂  Context & Current Design
─────────────────────────────
• Script purpose  :  Scan **~5,000 archived/suspended Google Workspace accounts**, find all mobile devices still linked, and (optionally) delete them.  
• Key entry points :  startScan(), scanDevicePreview(), getUserDevices() (Admin SDK → `Mobiledevices.list/remove`).  
• Core settings  :  `BATCH_SIZE = 100`, `MAX_EXECUTION_TIME = 4.5 min`, exponential back-off with 3 retries.  
• Processing style :  Strictly sequential—one user, one REST call at a time.  
• Result      :  Full run currently takes **hours**, often bumping into Apps Script 6-min execution limits and Admin SDK quotas.

─────────────────────────────
🎯  Optimization Goals
─────────────────────────────
Slash total scan time from **hours → minutes** while **never** exceeding Google API quotas or bricking devices, and while keeping error-handling robust.

1. **Batch API Queries**  
   • Investigate whether Admin SDK (Users: list, Directory: mobiledevices.list with `query=email:(user1 OR user2 …)` ) supports multi-user filtering.  

2. **Parallel Processing**  
   • Explore _concurrent_ UrlFetchApp / HTTP requests via Apps Script's **"parallel HTTP"** pattern or Cloud Functions/Workflows triggered from GAS.  
   • Evaluate trade-offs between native GAS triggers vs. off-loading to Cloud Run for fan-out/fan-in.

3. **API-Quota Optimization**  
   • Identify minimal projections/fields (e.g., `projection=BASIC`) required for preview mode.  
   • Fine-tune back-off & retry windows; convert fixed 60-second continuation to dynamic scheduling.

4. **Alternative Data Sources**  
   • Compare speed of Admin **Reports API (Mobile Details)** or **BigQuery exports** vs. per-user Mobiledevices.list.  
   • Evaluate _Customer Usage Reports_ or _Takeout transfer_ as bulk data shortcuts.

5. **Intelligent Caching**  
   • Design a stateful cache (PropertiesService / Firestore) holding "last scanned" timestamp & device hash; skip unchanged accounts.

─────────────────────────────
🔍  Research Deliverables
─────────────────────────────
Return a structured report with these sections:

A. **Executive Summary** – 1-2 paragraphs on the fastest, quota-safe path to the target scan time.  
B. **Feasibility Matrix** – Table scoring each optimization (effort vs. benefit vs. risk).  
C. **Deep Dive Findings** – Citations from Google official docs, known quotas (per-minute, per-user, per-customer), and any community benchmarks.  
D. **Prototype Snippets** – Idiomatic Apps Script or accompanying Cloud Run/Python examples that demonstrate the recommended approach (focus on batching & parallelism).  
E. **Implementation Roadmap** – Ordered steps, expected time savings, and test strategy (unit + load).  
F. **Fallback & Monitoring** – How to detect quota exhaustion early and auto-throttle; logging/alerting best practices.  

Use clear headings, bullet lists, and cite every claim with an **official Google source URL**. Where data is unavailable, propose an experiment to measure it.

Return your answer in **Markdown**.
```

## Implementation Priority

Based on expected impact vs. effort, research should focus on:

1. **High Impact, Low Effort**: API query optimization (projection, batch parameters)
2. **High Impact, Medium Effort**: Batch API queries for multiple users
3. **High Impact, High Effort**: Parallel processing architecture
4. **Medium Impact, Low Effort**: Intelligent caching for unchanged accounts
5. **Medium Impact, High Effort**: Alternative data sources (Reports API, BigQuery)

## Success Metrics

Target performance improvements:
- **Primary Goal**: Reduce 5,000 account scan from hours to minutes
- **Secondary Goal**: Maintain 100% reliability and error handling
- **Constraint**: Stay within all Google API quotas and limits

## Research Results

*This section will be populated with findings from the optimization research.*

## Future Benchmarks

Once optimizations are implemented, document:
- Before/after timing comparisons
- API call reduction percentages  
- Memory and execution time improvements
- Quota utilization analysis