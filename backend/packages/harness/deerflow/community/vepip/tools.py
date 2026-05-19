import json
import os

import httpx
from langchain.tools import tool

_CONVEX_SITE_URL = os.environ.get("VEPIP_CONVEX_SITE_URL", "")
_INTERNAL_SECRET = os.environ.get("VEPIP_INTERNAL_SECRET", "")


def _headers() -> dict:
    return {
        "Authorization": f"Bearer {_INTERNAL_SECRET}",
        "Content-Type": "application/json",
    }


def _post(path: str, payload: dict) -> dict:
    url = f"{_CONVEX_SITE_URL}{path}"
    with httpx.Client(timeout=30.0) as client:
        resp = client.post(url, json=payload, headers=_headers())
        resp.raise_for_status()
        return resp.json()


@tool("get_project_context", parse_docstring=True)
def get_project_context_tool(project_id: str, user_email: str) -> str:
    """Get full context for a VEPIP project: deliverables, budget categories, recent activities, milestones, and alerts.
    Always call this first when the user mentions a specific project.

    Args:
        project_id: The Convex ID of the project (e.g. 'j97abc123def').
        user_email: The email address of the requesting user.
    """
    data = _post("/ai/project-context", {"projectId": project_id, "userEmail": user_email})
    return json.dumps(data, indent=2)


@tool("log_activity", parse_docstring=True)
def log_activity_tool(
    project_id: str,
    user_email: str,
    title: str,
    activity_date: str,
    state: str = "",
    location: str = "",
    teachers_reached: int = 0,
    students_reached: int = 0,
    schools_reached: int = 0,
    notes: str = "",
    testimonial: str = "",
    testimonial_by: str = "",
) -> str:
    """Log a field activity for a VEPIP project. Only call this AFTER the user has confirmed the details.

    Args:
        project_id: The Convex ID of the project.
        user_email: The email address of the requesting user.
        title: Short descriptive title of the activity (e.g. 'Teacher training workshop in Mysore').
        activity_date: Date in YYYY-MM-DD format.
        state: Indian state name where the activity happened.
        location: Specific location or district.
        teachers_reached: Number of teachers reached (0 if not mentioned).
        students_reached: Number of students reached (0 if not mentioned).
        schools_reached: Number of schools reached (0 if not mentioned).
        notes: Any additional notes or observations.
        testimonial: A quote or testimonial from a beneficiary or participant.
        testimonial_by: Name of the person who gave the testimonial.
    """
    payload = {
        "projectId": project_id,
        "userEmail": user_email,
        "title": title,
        "activityDate": activity_date,
    }
    if state:
        payload["state"] = state
    if location:
        payload["location"] = location
    if teachers_reached:
        payload["teachersReached"] = teachers_reached
    if students_reached:
        payload["studentsReached"] = students_reached
    if schools_reached:
        payload["schoolsReached"] = schools_reached
    if notes:
        payload["notes"] = notes
    if testimonial:
        payload["testimonial"] = testimonial
    if testimonial_by:
        payload["testimonialBy"] = testimonial_by
    result = _post("/ai/log-activity", payload)
    return f"Activity logged successfully. ID: {result.get('id', 'unknown')}"


@tool("record_expense", parse_docstring=True)
def record_expense_tool(
    project_id: str,
    user_email: str,
    category_id: str,
    amount: float,
    description: str,
    spent_on: str,
    payment_mode: str = "",
) -> str:
    """Record an expense for a VEPIP project. Only call AFTER user confirms. Get category IDs from get_project_context first.

    Args:
        project_id: The Convex ID of the project.
        user_email: The email address of the requesting user.
        category_id: The Convex ID of the budget category (get from get_project_context).
        amount: The amount spent in rupees.
        description: Description of what was spent on.
        spent_on: Date in YYYY-MM-DD format.
        payment_mode: Payment method (cash, UPI, cheque, etc.).
    """
    payload = {
        "projectId": project_id,
        "userEmail": user_email,
        "categoryId": category_id,
        "amount": amount,
        "description": description,
        "spentOn": spent_on,
    }
    if payment_mode:
        payload["paymentMode"] = payment_mode
    result = _post("/ai/record-expense", payload)
    return f"Expense of ₹{amount} recorded. ID: {result.get('id', 'unknown')}"


@tool("update_deliverable", parse_docstring=True)
def update_deliverable_tool(
    user_email: str,
    deliverable_id: str,
    achieved: int,
) -> str:
    """Update the achieved count for a project deliverable. Only call AFTER user confirms. Get deliverable IDs from get_project_context.

    Args:
        user_email: The email address of the requesting user.
        deliverable_id: The Convex ID of the deliverable.
        achieved: The new total achieved count (not a delta — set the absolute total).
    """
    _post("/ai/update-deliverable", {"userEmail": user_email, "deliverableId": deliverable_id, "achieved": achieved})
    return f"Deliverable updated to {achieved} achieved."


@tool("add_milestone", parse_docstring=True)
def add_milestone_tool(
    user_email: str,
    project_id: str,
    title: str,
    due_date: str,
) -> str:
    """Add a new milestone to a project. Only call AFTER user confirms.

    Args:
        user_email: The email address of the requesting user.
        project_id: The Convex ID of the project.
        title: Title of the milestone.
        due_date: Due date in YYYY-MM-DD format.
    """
    result = _post("/ai/add-milestone", {"userEmail": user_email, "projectId": project_id, "title": title, "dueDate": due_date})
    return f"Milestone '{title}' added. ID: {result.get('id', 'unknown')}"


@tool("add_testimonial", parse_docstring=True)
def add_testimonial_tool(
    user_email: str,
    project_id: str,
    content: str,
    author: str,
    role: str = "",
) -> str:
    """Record an impact testimonial for a project. Only call AFTER user confirms.

    Args:
        user_email: The email address of the requesting user.
        project_id: The Convex ID of the project.
        content: The testimonial quote or story.
        author: Name of the person who gave the testimonial.
        role: Their role (e.g. 'Teacher', 'Parent', 'Student').
    """
    payload = {"userEmail": user_email, "projectId": project_id, "content": content, "author": author}
    if role:
        payload["role"] = role
    result = _post("/ai/add-testimonial", payload)
    return f"Testimonial from {author} recorded. ID: {result.get('id', 'unknown')}"


@tool("get_org_summary", parse_docstring=True)
def get_org_summary_tool(user_email: str) -> str:
    """Get organisation-wide summary: all active projects, total grant amount, at-risk count. Use for leadership/analytics questions.

    Args:
        user_email: The email address of the requesting user.
    """
    data = _post("/ai/org-summary", {"userEmail": user_email})
    return json.dumps(data, indent=2)


@tool("write_alert", parse_docstring=True)
def write_alert_tool(
    user_email: str,
    project_id: str,
    title: str,
    severity: str,
) -> str:
    """Create a proactive alert for a project. Use when you identify a risk during analysis.

    Args:
        user_email: The email address of the requesting user.
        project_id: The Convex ID of the project.
        title: Clear description of the alert (e.g. 'Budget for Travel >90% spent with 3 months remaining').
        severity: One of 'info', 'watch', or 'critical'.
    """
    result = _post("/ai/write-alert", {"userEmail": user_email, "projectId": project_id, "title": title, "severity": severity})
    return f"Alert created. ID: {result.get('id', 'unknown')}"


@tool("get_report_data", parse_docstring=True)
def get_report_data_tool(
    user_email: str,
    project_id: str,
    period_start: str,
    period_end: str,
) -> str:
    """Get all data needed to write a funder report: activities, expenses, deliverables, milestones for a period.

    Args:
        user_email: The email address of the requesting user.
        project_id: The Convex ID of the project.
        period_start: Start date in YYYY-MM-DD format.
        period_end: End date in YYYY-MM-DD format.
    """
    data = _post("/ai/report-data", {"userEmail": user_email, "projectId": project_id, "periodStart": period_start, "periodEnd": period_end})
    return json.dumps(data, indent=2)


@tool("query_portfolio", parse_docstring=True)
def query_portfolio_tool(
    user_email: str,
    theme: str = "",
    region: str = "",
    funder: str = "",
    from_date: str = "",
    to_date: str = "",
) -> str:
    """Search VEPIP's project portfolio by theme, region, or funder. Returns a
    list of matching projects with aggregate grant and spend totals.

    Use this when the user asks portfolio-level / cross-project questions
    that don't name a specific project — e.g. "show me all our Karnataka
    activities this quarter", "which projects are funded by Wipro", "what
    teacher-training work do we have running". Always prefer this over
    fetching every project with get_project_context one at a time.

    Args:
        user_email: The email address of the requesting user.
        theme: Optional theme keyword (e.g. "teacher training", "assistive tech").
        region: Optional Indian state name or two-letter code.
        funder: Optional funder name fragment.
        from_date: ISO date (YYYY-MM-DD) for project term overlap start.
        to_date: ISO date (YYYY-MM-DD) for project term overlap end.
    """
    payload: dict = {"userEmail": user_email}
    if theme: payload["theme"] = theme
    if region: payload["region"] = region
    if funder: payload["funder"] = funder
    if from_date: payload["fromDate"] = from_date
    if to_date: payload["toDate"] = to_date
    data = _post("/ai/query-portfolio", payload)
    return json.dumps(data, indent=2)


@tool("get_entity_profile", parse_docstring=True)
def get_entity_profile_tool(
    user_email: str,
    kind: str,
    entity_id: str,
) -> str:
    """Get the profile of a specific entity (funder, person, region, theme,
    school) including its rollup metrics, user-confirmed facts, and
    relationships to projects.

    Use this when a query_portfolio result mentions an entity and the user
    asks about it specifically.

    Args:
        user_email: The email address of the requesting user.
        kind: One of 'funder', 'person', 'region', 'theme', 'school'.
        entity_id: The Convex entity ID (returned by query_portfolio).
    """
    data = _post("/ai/entity-profile", {
        "userEmail": user_email, "kind": kind, "entityId": entity_id,
    })
    return json.dumps(data, indent=2)


@tool("remember_fact", parse_docstring=True)
def remember_fact_tool(
    user_email: str,
    entity_id: str,
    fact: str,
    confidence: float = 0.9,
) -> str:
    """Record a user-confirmed fact about an entity (funder, region, etc.) so
    that future conversations can recall it. The fact survives the nightly
    entity-graph rebuild. Use this only when the user explicitly tells you
    something durable about an entity ('Wipro prefers PDF reports',
    'Karnataka activities cluster around Mysore').

    Permissions: admin / leadership only. PMs cannot write facts.

    Args:
        user_email: The email address of the requesting user.
        entity_id: The Convex entity ID.
        fact: One-sentence fact to remember (<= 500 chars).
        confidence: 0..1, default 0.9 for user-confirmed.
    """
    data = _post("/ai/remember-fact", {
        "userEmail": user_email, "entityId": entity_id,
        "fact": fact, "confidence": confidence,
    })
    return f"Fact remembered (id={data.get('id')})"


@tool("search_knowledge", parse_docstring=True)
def search_knowledge_tool(
    user_email: str,
    query: str,
    project_id: str = "",
    kinds: str = "",
    date_from: str = "",
    date_to: str = "",
    top_k: int = 8,
) -> str:
    """Search VEPIP's knowledge base using semantic similarity. Returns ranked
    text chunks from project documents, past reports, activity narratives,
    testimonials, and uploaded MoUs/proposals.

    ALWAYS call this BEFORE answering any narrative question — "what did our
    last report say", "what was discussed in", "summarise the MoU", "find
    examples of X in our activities", "show me testimonials about Y". Never
    invent quoted text or specific narrative content; if this tool returns no
    results, say so plainly.

    Args:
        user_email: The email address of the requesting user.
        query: Natural-language search query.
        project_id: Optional Convex project ID to scope the search.
        kinds: Optional comma-separated list of document kinds to filter by
            (project_summary, mou, proposal, report_draft, activity_note,
            testimonial, meeting_note, uploaded_pdf).
        date_from: Optional ISO date (YYYY-MM-DD) — only documents created on
            or after this date.
        date_to: Optional ISO date (YYYY-MM-DD) — only documents created on or
            before this date.
        top_k: Number of results to return (default 8, max 25).
    """
    payload: dict = {"userEmail": user_email, "query": query, "topK": top_k}
    filters: dict = {}
    if project_id:
        filters["projectId"] = project_id
    if kinds:
        kind_list = [k.strip() for k in kinds.split(",") if k.strip()]
        if kind_list:
            filters["kinds"] = kind_list
    if date_from:
        filters["dateFrom"] = date_from
    if date_to:
        filters["dateTo"] = date_to
    if filters:
        payload["filters"] = filters
    data = _post("/ai/search-knowledge", payload)
    return json.dumps(data, indent=2)
