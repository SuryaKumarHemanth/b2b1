import zipfile
import io
import os
import csv
from datetime import datetime

# Helper function to convert column index to Excel column name (1 -> A, 27 -> AA, etc.)
def col_name(n):
    result = ""
    while n > 0:
        n, remainder = divmod(n - 1, 26)
        result = chr(65 + remainder) + result
    return result

def escape_xml(s):
    if s is None:
        return ""
    return str(s).replace("&", "&amp;").replace("<", "&lt;").replace(">", "&gt;").replace('"', "&quot;").replace("'", "&apos;")

# Comprehensive Dataset Definition
# 26 target accounts across Payroll, HR-Tech, Bank/BFSI
LEADS_DATA = [
    # Payroll Segment (9 accounts)
    {"id": "L001", "company": "RazorpayX Payroll", "segment": "Payroll", "contact": "Rahul Sharma", "title": "Head of Strategic Partnerships", "channel": "LinkedIn", "date_init": "2026-10-06", "stage": "Stage 4: Call Booked", "stage_num": 4, "value": 25000, "prob": 0.50, "status": "Replied - Advancing", "last_touch": "2026-10-21", "next_touch": "2026-10-28", "fu_status": "Scheduled", "notes": "Deeply interested in automated contractor tax deductions; intro call scheduled."},
    {"id": "L002", "company": "Keka HR & Payroll", "segment": "Payroll", "contact": "Ananya Sen", "title": "VP Strategic Alliances", "channel": "LinkedIn", "date_init": "2026-10-07", "stage": "Stage 4: Call Booked", "stage_num": 4, "value": 20000, "prob": 0.50, "status": "Replied - Advancing", "last_touch": "2026-10-22", "next_touch": "2026-10-29", "fu_status": "Scheduled", "notes": "Evaluating workflow add-ons for payroll processing; discovery call set."},
    {"id": "L003", "company": "greytHR", "segment": "Payroll", "contact": "Vikram Malhotra", "title": "Director of Product Alliances", "channel": "Email", "date_init": "2026-10-08", "stage": "Stage 5: Scoping", "stage_num": 5, "value": 30000, "prob": 0.70, "status": "Replied - Advancing", "last_touch": "2026-10-24", "next_touch": "2026-11-02", "fu_status": "Scheduled", "notes": "Completed discovery call. Technical architecture and API scoping in progress."},
    {"id": "L004", "company": "ZingHR", "segment": "Payroll", "contact": "Neha Deshmukh", "title": "Head of Global Alliances", "channel": "LinkedIn", "date_init": "2026-10-09", "stage": "Stage 6: Proposal Review", "stage_num": 6, "value": 35000, "prob": 0.85, "status": "Replied - Advancing", "last_touch": "2026-10-26", "next_touch": "2026-11-03", "fu_status": "Scheduled", "notes": "Advanced past technical review. Currently reviewing revenue-share and terms draft."},
    {"id": "L005", "company": "Rippling (India APAC)", "segment": "Payroll", "contact": "David Chen", "title": "Business Development Lead", "channel": "Email", "date_init": "2026-10-12", "stage": "Stage 2: Contacted", "stage_num": 2, "value": 28000, "prob": 0.10, "status": "No Reply", "last_touch": "2026-10-19", "next_touch": "2026-10-23", "fu_status": "Overdue (3d)", "notes": "Sent Touch 2 follow-up with APAC integration case study. Follow-up 3 overdue."},
    {"id": "L006", "company": "Deel (India Partnerships)", "segment": "Payroll", "contact": "Sarah Jenkins", "title": "Partnership Ecosystem Manager", "channel": "LinkedIn", "date_init": "2026-10-14", "stage": "Lost / Closed", "stage_num": 0, "value": 22000, "prob": 0.00, "status": "Not Interested", "last_touch": "2026-10-17", "next_touch": "-", "fu_status": "Closed", "notes": "Replied: Global product roadmap locked for H2; revisit in Q1 2027."},
    {"id": "L007", "company": "Pocket HRMS", "segment": "Payroll", "contact": "Amit Patel", "title": "Head of Product Alliances", "channel": "Email", "date_init": "2026-10-15", "stage": "Stage 2: Contacted", "stage_num": 2, "value": 15000, "prob": 0.10, "status": "No Reply", "last_touch": "2026-10-23", "next_touch": "2026-10-30", "fu_status": "Pending", "notes": "Touch 2 sent with client success metrics. Awaiting response."},
    {"id": "L008", "company": "TartanHQ", "segment": "Payroll", "contact": "Prateek Roy", "title": "Co-Founder & Head of Alliances", "channel": "LinkedIn", "date_init": "2026-10-16", "stage": "Stage 3: Replied", "stage_num": 3, "value": 18000, "prob": 0.25, "status": "Replied - Advancing", "last_touch": "2026-10-21", "next_touch": "2026-10-27", "fu_status": "Pending", "notes": "Replied positively regarding payroll API synergies; coordinating calendars for demo."},
    {"id": "L009", "company": "Payphi (PhiCommerce)", "segment": "Payroll", "contact": "Rajesh Pillai", "title": "VP Enterprise Alliances", "channel": "Email", "date_init": "2026-10-18", "stage": "Stage 2: Contacted", "stage_num": 2, "value": 16000, "prob": 0.10, "status": "No Reply", "last_touch": "2026-10-22", "next_touch": "2026-10-24", "fu_status": "Overdue (2d)", "notes": "No response to initial outreach; Touch 2 follow-up is overdue."},

    # HR-Tech Segment (10 accounts)
    {"id": "L010", "company": "Darwinbox", "segment": "HR-Tech", "contact": "Rohit Varma", "title": "VP Ecosystem & Marketplace", "channel": "LinkedIn", "date_init": "2026-10-06", "stage": "Stage 2: Contacted", "stage_num": 2, "value": 35000, "prob": 0.10, "status": "No Reply", "last_touch": "2026-10-18", "next_touch": "2026-10-23", "fu_status": "Overdue (3d)", "notes": "High priority account. Sent initial V1 pitch; follow-up overdue."},
    {"id": "L011", "company": "Zoho People", "segment": "HR-Tech", "contact": "Sridhar Narayanan", "title": "Director of Tech Partnerships", "channel": "Email", "date_init": "2026-10-07", "stage": "Stage 4: Call Booked", "stage_num": 4, "value": 24000, "prob": 0.50, "status": "Replied - Advancing", "last_touch": "2026-10-23", "next_touch": "2026-10-31", "fu_status": "Scheduled", "notes": "Replied to V2 case study. Scheduled partner screening call for end of month."},
    {"id": "L012", "company": "Sprinto Compliance", "segment": "HR-Tech", "contact": "Girish Reddi", "title": "Head of Ecosystem Growth", "channel": "LinkedIn", "date_init": "2026-10-09", "stage": "Stage 3: Replied", "stage_num": 3, "value": 22000, "prob": 0.25, "status": "Replied - Advancing", "last_touch": "2026-10-20", "next_touch": "2026-10-27", "fu_status": "Pending", "notes": "Expressed interest in continuous employee compliance integration; requested deck."},
    {"id": "L013", "company": "PeopleStrong", "segment": "HR-Tech", "contact": "Meenakshi Iyer", "title": "Director of Product Alliances", "channel": "Email", "date_init": "2026-10-10", "stage": "Stage 2: Contacted", "stage_num": 2, "value": 20000, "prob": 0.10, "status": "No Reply", "last_touch": "2026-10-19", "next_touch": "2026-10-26", "fu_status": "Pending", "notes": "Touch 2 sent with comparison matrix; waiting on reply."},
    {"id": "L014", "company": "Advantage Club", "segment": "HR-Tech", "contact": "Sourabh Deorah", "title": "Co-Founder & CEO", "channel": "LinkedIn", "date_init": "2026-10-11", "stage": "Stage 3: Replied", "stage_num": 3, "value": 15000, "prob": 0.25, "status": "Replied - Advancing", "last_touch": "2026-10-22", "next_touch": "2026-10-28", "fu_status": "Pending", "notes": "Connected on LinkedIn; replied to V2 case study asking for commercial structure."},
    {"id": "L015", "company": "Zaggle", "segment": "HR-Tech", "contact": "Avinash Godbole", "title": "VP Corporate Development", "channel": "Email", "date_init": "2026-10-13", "stage": "Stage 2: Contacted", "stage_num": 2, "value": 18000, "prob": 0.10, "status": "No Reply", "last_touch": "2026-10-20", "next_touch": "2026-10-28", "fu_status": "Pending", "notes": "Sent V1 pitch via email. Follow-up 1 sent on 10/20."},
    {"id": "L016", "company": "HONO HR", "segment": "HR-Tech", "contact": "Mukund Swaminathan", "title": "Chief Product Officer", "channel": "LinkedIn", "date_init": "2026-10-15", "stage": "Lost / Closed", "stage_num": 0, "value": 14000, "prob": 0.00, "status": "Not Interested", "last_touch": "2026-10-18", "next_touch": "-", "fu_status": "Closed", "notes": "Replied: Company currently undertaking core re-platforming; no bandwidth for APIs."},
    {"id": "L017", "company": "Zimyo", "segment": "HR-Tech", "contact": "Ajay Kumar", "title": "Head of Strategic Partnerships", "channel": "Email", "date_init": "2026-10-16", "stage": "Stage 2: Contacted", "stage_num": 2, "value": 12000, "prob": 0.10, "status": "No Reply", "last_touch": "2026-10-23", "next_touch": "2026-10-30", "fu_status": "Pending", "notes": "Touch 2 sent via email. No response yet."},
    {"id": "L018", "company": "Akrivia HCM", "segment": "HR-Tech", "contact": "Dinesh Rao", "title": "Strategic Alliances Lead", "channel": "LinkedIn", "date_init": "2026-10-17", "stage": "Stage 2: Contacted", "stage_num": 2, "value": 14000, "prob": 0.10, "status": "No Reply", "last_touch": "2026-10-22", "next_touch": "2026-10-29", "fu_status": "Pending", "notes": "LinkedIn InMail delivered; waiting on acceptance."},
    {"id": "L019", "company": "Workato (APAC HR)", "segment": "HR-Tech", "contact": "Kenji Takahashi", "title": "Director of Technology Alliances", "channel": "Email", "date_init": "2026-10-19", "stage": "Stage 2: Contacted", "stage_num": 2, "value": 25000, "prob": 0.10, "status": "No Reply", "last_touch": "2026-10-24", "next_touch": "2026-10-31", "fu_status": "Pending", "notes": "Sent V3 pitch. No response logged."},

    # Bank / BFSI Segment (7 accounts)
    {"id": "L020", "company": "HDFC Bank (SmartHub)", "segment": "Bank / BFSI", "contact": "Siddharth Mehta", "title": "VP Digital Alliances & Open Banking", "channel": "LinkedIn", "date_init": "2026-10-06", "stage": "Stage 2: Contacted", "stage_num": 2, "value": 50000, "prob": 0.10, "status": "No Reply", "last_touch": "2026-10-17", "next_touch": "2026-10-22", "fu_status": "Overdue (4d)", "notes": "High enterprise value target. Touch 2 follow-up 4 days overdue."},
    {"id": "L021", "company": "ICICI Bank (API Banking)", "segment": "Bank / BFSI", "contact": "Priya Sundaram", "title": "Head of Fintech Alliances & APIs", "channel": "Email", "date_init": "2026-10-08", "stage": "Stage 4: Call Booked", "stage_num": 4, "value": 45000, "prob": 0.50, "status": "Replied - Advancing", "last_touch": "2026-10-23", "next_touch": "2026-10-30", "fu_status": "Scheduled", "notes": "Replied to technical API pitch. Introduction call scheduled with Fintech onboarding team."},
    {"id": "L022", "company": "Axis Bank (Neo Banking)", "segment": "Bank / BFSI", "contact": "Sameer Joshi", "title": "VP Fintech Partnerships", "channel": "LinkedIn", "date_init": "2026-10-10", "stage": "Stage 2: Contacted", "stage_num": 2, "value": 40000, "prob": 0.10, "status": "No Reply", "last_touch": "2026-10-18", "next_touch": "2026-10-27", "fu_status": "Pending", "notes": "Sent V3 message via LinkedIn. Follow-up 1 sent on 10/18."},
    {"id": "L023", "company": "Yes Bank (Open Banking)", "segment": "Bank / BFSI", "contact": "Tarun Kapoor", "title": "Director of Corporate Alliances", "channel": "Email", "date_init": "2026-10-12", "stage": "Stage 2: Contacted", "stage_num": 2, "value": 30000, "prob": 0.10, "status": "No Reply", "last_touch": "2026-10-20", "next_touch": "2026-10-27", "fu_status": "Pending", "notes": "Email pitch delivered. Follow-up sent on 10/20."},
    {"id": "L024", "company": "IDFC FIRST Bank", "segment": "Bank / BFSI", "contact": "Pooja Nair", "title": "Lead - Strategic Integrations", "channel": "LinkedIn", "date_init": "2026-10-14", "stage": "Lost / Closed", "stage_num": 0, "value": 30000, "prob": 0.00, "status": "Not Interested", "last_touch": "2026-10-19", "next_touch": "-", "fu_status": "Closed", "notes": "Declined: Current API partnership sandbox closed to new aggregators until Q2."},
    {"id": "L025", "company": "M2P Fintech (Banking APIs)", "segment": "Bank / BFSI", "contact": "Karthik Sridhar", "title": "VP Strategic Partnerships", "channel": "LinkedIn", "date_init": "2026-10-16", "stage": "Stage 2: Contacted", "stage_num": 2, "value": 35000, "prob": 0.10, "status": "No Reply", "last_touch": "2026-10-21", "next_touch": "2026-10-28", "fu_status": "Pending", "notes": "V1 message delivered. Follow-up planned."},
    {"id": "L026", "company": "Karbon Card (BFSI Corporate)", "segment": "Bank / BFSI", "contact": "Sunil Sinha", "title": "Head of Strategic Growth", "channel": "Email", "date_init": "2026-10-18", "stage": "Lost / Closed", "stage_num": 0, "value": 20000, "prob": 0.00, "status": "Not Interested", "last_touch": "2026-10-21", "next_touch": "-", "fu_status": "Closed", "notes": "Replied: Focused solely on direct corporate sales; not pursuing partnership integrations."}
]

# Detailed Touch-by-Touch Outreach Log (54 logged outreach touches)
OUTREACH_LOG = [
    # L001 RazorpayX
    {"tid": "T001", "lid": "L001", "company": "RazorpayX Payroll", "contact": "Rahul Sharma", "touch": "Touch 1: Initial Pitch", "channel": "LinkedIn InMail", "msg_ver": "V1 - Short & Pain-Point", "sent_date": "2026-10-06", "outcome": "Replied", "resp_date": "2026-10-09", "days": 3, "notes": "Replied expressing interest in reducing contractor tax friction."},
    {"tid": "T002", "lid": "L001", "company": "RazorpayX Payroll", "contact": "Rahul Sharma", "touch": "Touch 2: Scheduling", "channel": "LinkedIn InMail", "msg_ver": "V1 - Short & Pain-Point", "sent_date": "2026-10-11", "outcome": "Replied", "resp_date": "2026-10-12", "days": 1, "notes": "Shared calendar link for partner discovery discussion."},
    {"tid": "T003", "lid": "L001", "company": "RazorpayX Payroll", "contact": "Rahul Sharma", "touch": "Touch 3: Confirmation", "channel": "Email - Direct", "msg_ver": "V1 - Short & Pain-Point", "sent_date": "2026-10-21", "outcome": "Call Booked", "resp_date": "2026-10-21", "days": 0, "notes": "Confirmed 30-min discovery call for 10/28."},

    # L002 Keka HR
    {"tid": "T004", "lid": "L002", "company": "Keka HR & Payroll", "contact": "Ananya Sen", "touch": "Touch 1: Initial Pitch", "channel": "LinkedIn Connection + Note", "msg_ver": "V1 - Short & Pain-Point", "sent_date": "2026-10-07", "outcome": "Replied", "resp_date": "2026-10-11", "days": 4, "notes": "Connected and replied: Interested in seeing integration architecture."},
    {"tid": "T005", "lid": "L002", "company": "Keka HR & Payroll", "contact": "Ananya Sen", "touch": "Touch 2: Follow-up Call", "channel": "LinkedIn InMail", "msg_ver": "V1 - Short & Pain-Point", "sent_date": "2026-10-14", "outcome": "Replied", "resp_date": "2026-10-16", "days": 2, "notes": "Proposed discovery sync with tech alliances team."},
    {"tid": "T006", "lid": "L002", "company": "Keka HR & Payroll", "contact": "Ananya Sen", "touch": "Touch 3: Agenda Lock", "channel": "Email - Direct", "msg_ver": "V1 - Short & Pain-Point", "sent_date": "2026-10-22", "outcome": "Call Booked", "resp_date": "2026-10-22", "days": 0, "notes": "Discovery call confirmed on calendar for 10/29."},

    # L003 greytHR
    {"tid": "T007", "lid": "L003", "company": "greytHR", "contact": "Vikram Malhotra", "touch": "Touch 1: Initial Pitch", "channel": "Email - Direct", "msg_ver": "V2 - Social Proof & Case Study", "sent_date": "2026-10-08", "outcome": "Replied", "resp_date": "2026-10-13", "days": 5, "notes": "Replied to case study. Asked for technical integration sandbox details."},
    {"tid": "T008", "lid": "L003", "company": "greytHR", "contact": "Vikram Malhotra", "touch": "Touch 2: Discovery Call", "channel": "Email - Direct", "msg_ver": "V2 - Social Proof & Case Study", "sent_date": "2026-10-15", "outcome": "Call Booked", "resp_date": "2026-10-16", "days": 1, "notes": "Completed initial call on 10/18; team advanced opportunity to tech evaluation."},
    {"tid": "T009", "lid": "L003", "company": "greytHR", "contact": "Vikram Malhotra", "touch": "Touch 3: Scoping Docs", "channel": "Email - Direct", "msg_ver": "V2 - Social Proof & Case Study", "sent_date": "2026-10-24", "outcome": "Replied", "resp_date": "2026-10-25", "days": 1, "notes": "greytHR architecture team reviewing API documentation."},

    # L004 ZingHR
    {"tid": "T010", "lid": "L004", "company": "ZingHR", "contact": "Neha Deshmukh", "touch": "Touch 1: Initial Pitch", "channel": "LinkedIn InMail", "msg_ver": "V2 - Social Proof & Case Study", "sent_date": "2026-10-09", "outcome": "Replied", "resp_date": "2026-10-12", "days": 3, "notes": "Interested in white-labeled add-on for enterprise tier."},
    {"tid": "T011", "lid": "L004", "company": "ZingHR", "contact": "Neha Deshmukh", "touch": "Touch 2: Discovery Call", "channel": "Email - Direct", "msg_ver": "V2 - Social Proof & Case Study", "sent_date": "2026-10-14", "outcome": "Call Booked", "resp_date": "2026-10-15", "days": 1, "notes": "Held discovery session on 10/19. Very strong product fit."},
    {"tid": "T012", "lid": "L004", "company": "ZingHR", "contact": "Neha Deshmukh", "touch": "Touch 3: Proposal Submission", "channel": "Email - Direct", "msg_ver": "V2 - Social Proof & Case Study", "sent_date": "2026-10-26", "outcome": "Replied", "resp_date": "2026-10-26", "days": 0, "notes": "Submitted commercial proposal ($35k). Terms in executive review."},

    # L005 Rippling
    {"tid": "T013", "lid": "L005", "company": "Rippling (India APAC)", "contact": "David Chen", "touch": "Touch 1: Initial Pitch", "channel": "Email - Direct", "msg_ver": "V1 - Short & Pain-Point", "sent_date": "2026-10-12", "outcome": "No Reply", "resp_date": "-", "days": "-", "notes": "Initial email delivered; no response after 7 days."},
    {"tid": "T014", "lid": "L005", "company": "Rippling (India APAC)", "contact": "David Chen", "touch": "Touch 2: Follow-up 1", "channel": "Email - Follow-up", "msg_ver": "V1 - Short & Pain-Point", "sent_date": "2026-10-19", "outcome": "No Reply", "resp_date": "-", "days": "-", "notes": "Sent APAC benchmark data. Follow-up 3 is currently 3 days overdue."},

    # L006 Deel
    {"tid": "T015", "lid": "L006", "company": "Deel (India Partnerships)", "contact": "Sarah Jenkins", "touch": "Touch 1: Initial Pitch", "channel": "LinkedIn InMail", "msg_ver": "V3 - Feature-Heavy Pitch", "sent_date": "2026-10-14", "outcome": "Replied", "resp_date": "2026-10-17", "days": 3, "notes": "Replied with polite rejection: Product roadmap committed for H2. Revisit in 2027."},
    {"tid": "T016", "lid": "L006", "company": "Deel (India Partnerships)", "contact": "Sarah Jenkins", "touch": "Touch 2: Acknowledgement", "channel": "LinkedIn InMail", "msg_ver": "V3 - Feature-Heavy Pitch", "sent_date": "2026-10-17", "outcome": "Not Interested", "resp_date": "2026-10-17", "days": 0, "notes": "Marked as closed / lost for current cycle."},

    # L007 Pocket HRMS
    {"tid": "T017", "lid": "L007", "company": "Pocket HRMS", "contact": "Amit Patel", "touch": "Touch 1: Initial Pitch", "channel": "Email - Direct", "msg_ver": "V1 - Short & Pain-Point", "sent_date": "2026-10-15", "outcome": "No Reply", "resp_date": "-", "days": "-", "notes": "Initial email sent."},
    {"tid": "T018", "lid": "L007", "company": "Pocket HRMS", "contact": "Amit Patel", "touch": "Touch 2: Follow-up 1", "channel": "Email - Follow-up", "msg_ver": "V1 - Short & Pain-Point", "sent_date": "2026-10-23", "outcome": "No Reply", "resp_date": "-", "days": "-", "notes": "Follow-up email sent with implementation timeline. Next check scheduled."},

    # L008 TartanHQ
    {"tid": "T019", "lid": "L008", "company": "TartanHQ", "contact": "Prateek Roy", "touch": "Touch 1: Initial Pitch", "channel": "LinkedIn InMail", "msg_ver": "V2 - Social Proof & Case Study", "sent_date": "2026-10-16", "outcome": "Replied", "resp_date": "2026-10-21", "days": 5, "notes": "Replied: Keen to explore API data aggregation collaboration."},
    {"tid": "T020", "lid": "L008", "company": "TartanHQ", "contact": "Prateek Roy", "touch": "Touch 2: Follow-up Demo", "channel": "LinkedIn InMail", "msg_ver": "V2 - Social Proof & Case Study", "sent_date": "2026-10-23", "outcome": "Replied", "resp_date": "2026-10-25", "days": 2, "notes": "Coordinating availability for executive demo."},

    # L009 Payphi
    {"tid": "T021", "lid": "L009", "company": "Payphi (PhiCommerce)", "contact": "Rajesh Pillai", "touch": "Touch 1: Initial Pitch", "channel": "Email - Direct", "msg_ver": "V3 - Feature-Heavy Pitch", "sent_date": "2026-10-18", "outcome": "No Reply", "resp_date": "-", "days": "-", "notes": "Long feature-heavy email sent; no open/click response."},
    {"tid": "T022", "lid": "L009", "company": "Payphi (PhiCommerce)", "contact": "Rajesh Pillai", "touch": "Touch 2: Follow-up 1", "channel": "Email - Follow-up", "msg_ver": "V3 - Feature-Heavy Pitch", "sent_date": "2026-10-22", "outcome": "No Reply", "resp_date": "-", "days": "-", "notes": "Follow-up email sent. 2 days overdue for Touch 3."},

    # L010 Darwinbox
    {"tid": "T023", "lid": "L010", "company": "Darwinbox", "contact": "Rohit Varma", "touch": "Touch 1: Initial Pitch", "channel": "LinkedIn InMail", "msg_ver": "V1 - Short & Pain-Point", "sent_date": "2026-10-06", "outcome": "No Reply", "resp_date": "-", "days": "-", "notes": "Sent short pain-point message to VP Ecosystem."},
    {"tid": "T024", "lid": "L010", "company": "Darwinbox", "contact": "Rohit Varma", "touch": "Touch 2: Follow-up 1", "channel": "LinkedIn InMail", "msg_ver": "V1 - Short & Pain-Point", "sent_date": "2026-10-18", "outcome": "No Reply", "resp_date": "-", "days": "-", "notes": "Second ping sent. Touch 3 follow-up is currently 3 days overdue."},

    # L011 Zoho People
    {"tid": "T025", "lid": "L011", "company": "Zoho People", "contact": "Sridhar Narayanan", "touch": "Touch 1: Initial Pitch", "channel": "Email - Direct", "msg_ver": "V2 - Social Proof & Case Study", "sent_date": "2026-10-07", "outcome": "No Reply", "resp_date": "-", "days": "-", "notes": "Email sent with HR-tech integration case study."},
    {"tid": "T026", "lid": "L011", "company": "Zoho People", "contact": "Sridhar Narayanan", "touch": "Touch 2: Follow-up 1", "channel": "Email - Follow-up", "msg_ver": "V2 - Social Proof & Case Study", "sent_date": "2026-10-16", "outcome": "Replied", "resp_date": "2026-10-20", "days": 4, "notes": "Replied: Forwarded to Developer Marketplace team for review."},
    {"tid": "T027", "lid": "L011", "company": "Zoho People", "contact": "Sridhar Narayanan", "touch": "Touch 3: Screening Call", "channel": "Email - Direct", "msg_ver": "V2 - Social Proof & Case Study", "sent_date": "2026-10-23", "outcome": "Call Booked", "resp_date": "2026-10-24", "days": 1, "notes": "Partner screening call locked in for 10/31."},

    # L012 Sprinto
    {"tid": "T028", "lid": "L012", "company": "Sprinto Compliance", "contact": "Girish Reddi", "touch": "Touch 1: Initial Pitch", "channel": "LinkedIn InMail", "msg_ver": "V1 - Short & Pain-Point", "sent_date": "2026-10-09", "outcome": "Replied", "resp_date": "2026-10-15", "days": 6, "notes": "Replied: Highly interested in payroll security compliance hooks."},
    {"tid": "T029", "lid": "L012", "company": "Sprinto Compliance", "contact": "Girish Reddi", "touch": "Touch 2: Architecture Share", "channel": "LinkedIn InMail", "msg_ver": "V1 - Short & Pain-Point", "sent_date": "2026-10-20", "outcome": "Replied", "resp_date": "2026-10-22", "days": 2, "notes": "Shared 2-page integration architecture overview."},

    # L013 PeopleStrong
    {"tid": "T030", "lid": "L013", "company": "PeopleStrong", "contact": "Meenakshi Iyer", "touch": "Touch 1: Initial Pitch", "channel": "Email - Direct", "msg_ver": "V3 - Feature-Heavy Pitch", "sent_date": "2026-10-10", "outcome": "No Reply", "resp_date": "-", "days": "-", "notes": "Sent V3 pitch. No response."},
    {"tid": "T031", "lid": "L013", "company": "PeopleStrong", "contact": "Meenakshi Iyer", "touch": "Touch 2: Follow-up 1", "channel": "Email - Follow-up", "msg_ver": "V3 - Feature-Heavy Pitch", "sent_date": "2026-10-19", "outcome": "No Reply", "resp_date": "-", "days": "-", "notes": "Sent comparison matrix; still no reply."},

    # L014 Advantage Club
    {"tid": "T032", "lid": "L014", "company": "Advantage Club", "contact": "Sourabh Deorah", "touch": "Touch 1: Initial Pitch", "channel": "LinkedIn Connection + Note", "msg_ver": "V2 - Social Proof & Case Study", "sent_date": "2026-10-11", "outcome": "Replied", "resp_date": "2026-10-17", "days": 6, "notes": "Connected and replied: Interested in corporate perks cross-distribution."},
    {"tid": "T033", "lid": "L014", "company": "Advantage Club", "contact": "Sourabh Deorah", "touch": "Touch 2: Commercial Sync", "channel": "LinkedIn InMail", "msg_ver": "V2 - Social Proof & Case Study", "sent_date": "2026-10-22", "outcome": "Replied", "resp_date": "2026-10-24", "days": 2, "notes": "Sent commercial outline; evaluating next steps."},

    # L015 Zaggle
    {"tid": "T034", "lid": "L015", "company": "Zaggle", "contact": "Avinash Godbole", "touch": "Touch 1: Initial Pitch", "channel": "Email - Direct", "msg_ver": "V1 - Short & Pain-Point", "sent_date": "2026-10-13", "outcome": "No Reply", "resp_date": "-", "days": "-", "notes": "Direct email sent to VP Corp Dev."},
    {"tid": "T035", "lid": "L015", "company": "Zaggle", "contact": "Avinash Godbole", "touch": "Touch 2: Follow-up 1", "channel": "Email - Follow-up", "msg_ver": "V1 - Short & Pain-Point", "sent_date": "2026-10-20", "outcome": "No Reply", "resp_date": "-", "days": "-", "notes": "Follow-up sent; no response recorded."},

    # L016 HONO HR
    {"tid": "T036", "lid": "L016", "company": "HONO HR", "contact": "Mukund Swaminathan", "touch": "Touch 1: Initial Pitch", "channel": "LinkedIn InMail", "msg_ver": "V3 - Feature-Heavy Pitch", "sent_date": "2026-10-15", "outcome": "Replied", "resp_date": "2026-10-18", "days": 3, "notes": "Replied stating that internal product bandwidth is saturated."},
    {"tid": "T037", "lid": "L016", "company": "HONO HR", "contact": "Mukund Swaminathan", "touch": "Touch 2: Wrap-up", "channel": "LinkedIn InMail", "msg_ver": "V3 - Feature-Heavy Pitch", "sent_date": "2026-10-18", "outcome": "Not Interested", "resp_date": "2026-10-18", "days": 0, "notes": "Logged as not interested / bad timing."},

    # L017 Zimyo
    {"tid": "T038", "lid": "L017", "company": "Zimyo", "contact": "Ajay Kumar", "touch": "Touch 1: Initial Pitch", "channel": "Email - Direct", "msg_ver": "V2 - Social Proof & Case Study", "sent_date": "2026-10-16", "outcome": "No Reply", "resp_date": "-", "days": "-", "notes": "Email sent with SMB payroll integration numbers."},
    {"tid": "T039", "lid": "L017", "company": "Zimyo", "contact": "Ajay Kumar", "touch": "Touch 2: Follow-up 1", "channel": "Email - Follow-up", "msg_ver": "V2 - Social Proof & Case Study", "sent_date": "2026-10-23", "outcome": "No Reply", "resp_date": "-", "days": "-", "notes": "Follow-up email sent. Pending reply."},

    # L018 Akrivia HCM
    {"tid": "T040", "lid": "L018", "company": "Akrivia HCM", "contact": "Dinesh Rao", "touch": "Touch 1: Initial Pitch", "channel": "LinkedIn InMail", "msg_ver": "V1 - Short & Pain-Point", "sent_date": "2026-10-17", "outcome": "No Reply", "resp_date": "-", "days": "-", "notes": "InMail sent; no reply after 5 days."},
    {"tid": "T041", "lid": "L018", "company": "Akrivia HCM", "contact": "Dinesh Rao", "touch": "Touch 2: Follow-up 1", "channel": "LinkedIn InMail", "msg_ver": "V1 - Short & Pain-Point", "sent_date": "2026-10-22", "outcome": "No Reply", "resp_date": "-", "days": "-", "notes": "Gentle nudge sent on LinkedIn."},

    # L019 Workato
    {"tid": "T042", "lid": "L019", "company": "Workato (APAC HR)", "contact": "Kenji Takahashi", "touch": "Touch 1: Initial Pitch", "channel": "Email - Direct", "msg_ver": "V3 - Feature-Heavy Pitch", "sent_date": "2026-10-19", "outcome": "No Reply", "resp_date": "-", "days": "-", "notes": "Complex feature-heavy connector pitch sent; no reply."},
    {"tid": "T043", "lid": "L019", "company": "Workato (APAC HR)", "contact": "Kenji Takahashi", "touch": "Touch 2: Follow-up 1", "channel": "Email - Follow-up", "msg_ver": "V3 - Feature-Heavy Pitch", "sent_date": "2026-10-24", "outcome": "No Reply", "resp_date": "-", "days": "-", "notes": "Follow-up email delivered. Pending reply."},

    # L020 HDFC Bank
    {"tid": "T044", "lid": "L020", "company": "HDFC Bank (SmartHub)", "contact": "Siddharth Mehta", "touch": "Touch 1: Initial Pitch", "channel": "LinkedIn InMail", "msg_ver": "V1 - Short & Pain-Point", "sent_date": "2026-10-06", "outcome": "No Reply", "resp_date": "-", "days": "-", "notes": "Initial pitch sent via LinkedIn InMail."},
    {"tid": "T045", "lid": "L020", "company": "HDFC Bank (SmartHub)", "contact": "Siddharth Mehta", "touch": "Touch 2: Follow-up 1", "channel": "LinkedIn InMail", "msg_ver": "V1 - Short & Pain-Point", "sent_date": "2026-10-17", "outcome": "No Reply", "resp_date": "-", "days": "-", "notes": "Follow-up sent. Touch 3 is critically overdue (4 days)."},

    # L021 ICICI Bank
    {"tid": "T046", "lid": "L021", "company": "ICICI Bank (API Banking)", "contact": "Priya Sundaram", "touch": "Touch 1: Initial Pitch", "channel": "Email - Direct", "msg_ver": "V2 - Social Proof & Case Study", "sent_date": "2026-10-08", "outcome": "No Reply", "resp_date": "-", "days": "-", "notes": "Email sent to API Banking partnerships inbox."},
    {"tid": "T047", "lid": "L021", "company": "ICICI Bank (API Banking)", "contact": "Priya Sundaram", "touch": "Touch 2: Technical Follow-up", "channel": "Email - Follow-up", "msg_ver": "V2 - Social Proof & Case Study", "sent_date": "2026-10-16", "outcome": "Replied", "resp_date": "2026-10-21", "days": 5, "notes": "Replied: Forwarded to Open Banking team for evaluation."},
    {"tid": "T048", "lid": "L021", "company": "ICICI Bank (API Banking)", "contact": "Priya Sundaram", "touch": "Touch 3: Intro Call", "channel": "Email - Direct", "msg_ver": "V2 - Social Proof & Case Study", "sent_date": "2026-10-23", "outcome": "Call Booked", "resp_date": "2026-10-24", "days": 1, "notes": "Introductory partner call scheduled for 10/30."},

    # L022 Axis Bank
    {"tid": "T049", "lid": "L022", "company": "Axis Bank (Neo Banking)", "contact": "Sameer Joshi", "touch": "Touch 1: Initial Pitch", "channel": "LinkedIn InMail", "msg_ver": "V3 - Feature-Heavy Pitch", "sent_date": "2026-10-10", "outcome": "No Reply", "resp_date": "-", "days": "-", "notes": "V3 feature pitch sent to VP Fintech Partnerships."},
    {"tid": "T050", "lid": "L022", "company": "Axis Bank (Neo Banking)", "contact": "Sameer Joshi", "touch": "Touch 2: Follow-up 1", "channel": "LinkedIn InMail", "msg_ver": "V3 - Feature-Heavy Pitch", "sent_date": "2026-10-18", "outcome": "No Reply", "resp_date": "-", "days": "-", "notes": "Follow-up sent; no response received."},

    # L023 Yes Bank
    {"tid": "T051", "lid": "L023", "company": "Yes Bank (Open Banking)", "contact": "Tarun Kapoor", "touch": "Touch 1: Initial Pitch", "channel": "Email - Direct", "msg_ver": "V1 - Short & Pain-Point", "sent_date": "2026-10-12", "outcome": "No Reply", "resp_date": "-", "days": "-", "notes": "Sent V1 pitch via corporate email."},
    {"tid": "T052", "lid": "L023", "company": "Yes Bank (Open Banking)", "contact": "Tarun Kapoor", "touch": "Touch 2: Follow-up 1", "channel": "Email - Follow-up", "msg_ver": "V1 - Short & Pain-Point", "sent_date": "2026-10-20", "outcome": "No Reply", "resp_date": "-", "days": "-", "notes": "Follow-up sent; no response."},

    # L024 IDFC FIRST Bank
    {"tid": "T053", "lid": "L024", "company": "IDFC FIRST Bank", "contact": "Pooja Nair", "touch": "Touch 1: Initial Pitch", "channel": "LinkedIn InMail", "msg_ver": "V2 - Social Proof & Case Study", "sent_date": "2026-10-14", "outcome": "Replied", "resp_date": "2026-10-19", "days": 5, "notes": "Replied: API gateway onboarding currently paused for aggregators."},

    # L025 M2P Fintech
    {"tid": "T054", "lid": "L025", "company": "M2P Fintech (Banking APIs)", "contact": "Karthik Sridhar", "touch": "Touch 1: Initial Pitch", "channel": "LinkedIn InMail", "msg_ver": "V1 - Short & Pain-Point", "sent_date": "2026-10-16", "outcome": "No Reply", "resp_date": "-", "days": "-", "notes": "V1 message delivered. Follow-up planned."},

    # L026 Karbon Card
    {"tid": "T055", "lid": "L026", "company": "Karbon Card (BFSI Corporate)", "contact": "Sunil Sinha", "touch": "Touch 1: Initial Pitch", "channel": "Email - Direct", "msg_ver": "V3 - Feature-Heavy Pitch", "sent_date": "2026-10-18", "outcome": "Replied", "resp_date": "2026-10-21", "days": 3, "notes": "Replied: Company is focusing 100% on direct sales; no partner channel."}
]

print(f"Dataset compiled: {len(LEADS_DATA)} leads, {len(OUTREACH_LOG)} outreach touches.")
