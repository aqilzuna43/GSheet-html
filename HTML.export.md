# Product Requirements Document (PRD)
## Visual Schedule Dashboard (Pilot)

### Version
v0.1 (Pilot)

---

## 1. Objective

Build a lightweight, interactive visual schedule dashboard that reads structured data from a Google Sheet and displays it in a calendar and table view, embedded inside Google Sites.

This is a visual layer only.  
All data entry and updates remain manual in Google Sheets.

---

## 2. Scope (Pilot Only)

### In Scope
- Google Sheet as single data source
- Interactive calendar view (Month / Week)
- Interactive table view
- Filtering and search
- Embedded inside Google Sites
- Domain-restricted access

### Out of Scope
- Editing inside dashboard
- Drag-and-drop scheduling
- Notifications
- Workflow automation
- Role-based permissions
- Integration with external systems

---

## 3. Target Users

- Engineering team members (5–20 users)
- Managers needing visual overview

---

## 4. Architecture

Google Sheet (Manual Input)
→ Google Apps Script (Web App)
→ HTML + JS UI (Calendar + Table)
→ Embedded in Google Sites

The Google Sheet remains the source of truth.

---

## 5. Functional Requirements

### 5.1 Data Source
Sheet Name: `Schedule`

Required Columns:
- ID
- Title
- Start Date
- End Date
- Owner
- Department
- Status
- Description (optional)
- Tags (optional)

Validation:
- End Date must be after Start Date
- Header row must not be modified

---

### 5.2 Calendar View
- Month view
- Week view
- Click event → show detail modal
- Color coding by Department or Status

---

### 5.3 Table View
- Sortable columns
- Global search
- Filter by:
  - Owner
  - Department
  - Status
- Pagination

---

### 5.4 Sync Behavior
- Reflect latest Sheet data on page refresh
- Optional manual “Refresh” button
- No caching older than 5 minutes

---

## 6. Non-Functional Requirements

- Load time < 3 seconds for 500 entries
- Support up to 1000 schedule rows
- Responsive on desktop and mobile
- Accessible only to internal domain users

---

## 7. Success Criteria (Pilot)

- Users can visually understand workload at a glance
- No need to open raw Sheet for viewing
- Stable for 2 weeks without schema break
- Positive feedback from pilot group

---

## 8. Risks

- Schema modification breaking UI
- Apps Script quota limitations
- Overbuilding beyond pilot needs

---

## 9. Pilot Exit Criteria

After 2 weeks:
- Collect feedback
- Identify improvement areas
- Decide:
  - Keep simple
  - Expand features
  - Or discontinue

---

## Principle

This system is a visualization layer, not a workflow tool.

Keep it simple.