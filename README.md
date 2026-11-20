# AI Lead Capture & Response Automation

A no-server automation that captures leads from a Google Form, writes a personalised
email reply with AI, sends it, and logs everything in Google Sheets.

## Flow
Google Form → Google Sheet → Apps Script (on form submit) → Groq API → Gmail reply → `Leads_Log` sheet

## Features
- AI-personalised reply based on the lead's service, budget and message
- Hot / Warm / Cold lead scoring
- Hot-lead alert email to the owner
- Duplicate detection (same email within 24 hours)
- Junk-message guard ("tell us more" email instead of an AI reply)
- AI-failure fallback template, so no lead is dropped
- Status column for every outcome: `SENT`, `SENT_NEEDS_DETAILS`, `SENT_FALLBACK`, `DUPLICATE_SKIPPED`, `ERROR`

## Tech stack
Google Forms, Google Sheets, Google Apps Script, Groq API (`openai/gpt-oss-120b`), Gmail

## Setup
1. Create a Google Form with these questions: Name, Email, Phone, Service Interested In, Budget, Message.
2. Link the form to a Google Sheet (Responses → Link to Sheets).
3. In the Sheet, open Extensions → Apps Script and paste `Code.gs`.
4. Project Settings → Script Properties → add `GROQ_API_KEY` with your key from console.groq.com.
5. Run `installTrigger` once and approve the permissions.
6. Submit the live form to test. Results appear in the `Leads_Log` tab.

Edit the `CONFIG` block at the top of `Code.gs` to change the business name, pitch, sender name or model.

## Roadmap
Automated follow-ups, WhatsApp replies (WhatsApp Cloud API), and a CRM pipeline view.

## Demo
<link to your video>