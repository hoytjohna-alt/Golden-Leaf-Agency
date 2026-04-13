# Agency Operating System

This app is now set up for:

- rep/producer logins
- rep-owned lead pipelines
- admin-wide agency oversight
- Render deployment with environment variables

## Stack

- Frontend: Vite + vanilla JavaScript
- Auth and database: Supabase
- Hosting: Render static site

## Core files

- `index.html` - app shell
- `app.js` - authenticated app logic
- `styles.css` - visual styling
- `supabase-schema.sql` - database tables, policies, and trigger
- `.env.example` - local env variable template
- `package.json` - tiny build/deploy setup
- `render.yaml` - Render blueprint config

## Local development

1. Copy `.env.example` to `.env`
2. Fill in:
   - `VITE_SUPABASE_URL`
   - `VITE_SUPABASE_PUBLISHABLE_KEY`
   - `VITE_APP_URL`
3. Install dependencies:

```bash
npm install
```

4. Run the dev server:

```bash
npm run dev
```

Vite will print a localhost URL.

## Supabase setup

### 1. Run the schema

Open Supabase SQL Editor and run:

- `supabase-schema.sql`

This creates:

- `profiles`
- `app_settings`
- `opportunities`
- `coaching_notes`
- row-level security policies
- an auth trigger that auto-creates a profile row for each new user

### 2. Create users

In Supabase Authentication:

- create the owner/admin user
- create each rep/producer user

Promote the owner to admin:

```sql
update public.profiles
set role = 'admin'
where email = 'hoyt@independentagencyconsulting.com';
```

### 3. Use the correct frontend key

Use:

- `Project URL`
- `Publishable key`

Do not use the secret key in the frontend.

## Render deployment

### 1. Create a new Static Site in Render

1. In Render, click `New +`
2. Choose `Static Site`
3. Connect the GitHub repository you just published

### 2. Build settings

Render should pick these up from `render.yaml`:

- Build Command: `npm install && npm run build`
- Publish Directory: `dist`

### 3. Add environment variables

In Render, add:

- `VITE_SUPABASE_URL`
- `VITE_SUPABASE_PUBLISHABLE_KEY`
- `VITE_APP_URL`

Use the same values from Supabase.

### 4. Deploy

Click deploy. Render will build the app and publish it online.

## Security model

The SQL schema enforces:

- reps can only read and edit their own opportunities
- admins can access all opportunities
- admins can manage global settings
- reps can only see their own profile
- admins can manage role and active status for user profiles

## Pre-client handoff

Before showing the platform to a client or rolling it into a live agency, use:

- [PRE_CLIENT_HANDOFF.md](/Users/johnhoyt/Desktop/Golden%20Leaf%20Agency/PRE_CLIENT_HANDOFF.md)

That checklist covers:

- admin and producer workflow QA
- renewal and reporting checks
- Claude and help-center checks
- mobile review
- demo readiness

## Current admin capabilities

- full agency dashboard
- cross-rep opportunity visibility
- rep assignment on leads
- coaching note management
- global assumptions and carrier commission edits
- profile role and active-status management inside the app
- password reset request from the login screen

## Claude assistant setup

The app now includes a floating Claude chat assistant for logged-in users.

Security model:

- the Anthropic API key stays server-side
- reps only get data they can already access through Supabase RLS
- admins can ask agency-wide questions

### 1. Set Supabase function secrets

In Supabase, add these secrets for Edge Functions:

- `ANTHROPIC_API_KEY`
- `ANTHROPIC_MODEL`

Suggested model:

- `claude-sonnet-4-20250514`

### 2. Deploy the Edge Function

Deploy:

- `supabase/functions/claude-assistant/index.ts`

Function name:

- `claude-assistant`

### 3. Keep frontend env vars the same

The frontend still only needs:

- `VITE_SUPABASE_URL`
- `VITE_SUPABASE_PUBLISHABLE_KEY`
- `VITE_APP_URL`

No Anthropic key should ever be added to Render frontend env vars.

### 4. Use the assistant

- admins can ask about agency metrics, renewals, pipeline, producers, and lead activity
- reps can ask only about their own scoped data
- if a lead is open in the workspace, the assistant also gets that lead as focused context

## Google Calendar setup

The app now includes Google Calendar connection and lead-to-calendar event creation.

### What it does

- each user connects their own Google Calendar
- reps can create calendar events from lead follow-ups and appointments
- admins can connect their own calendar separately
- Outlook is the next provider to add

### Google Cloud setup

Create OAuth credentials in Google Cloud and add this redirect URI:

- `https://wnmspwioshucrcqbfabo.supabase.co/functions/v1/google-calendar-callback`

Recommended scope used by the app:

- `https://www.googleapis.com/auth/calendar.events`

### Supabase secrets for calendar

Add these secrets in Supabase Edge Functions:

- `SUPABASE_SERVICE_ROLE_KEY`
- `GOOGLE_CLIENT_ID`
- `GOOGLE_CLIENT_SECRET`
- `APP_URL`

`APP_URL` should be your live Render URL, for example:

- `https://your-app-name.onrender.com`

### Functions to deploy

- `google-calendar-connect`
- `google-calendar-status`
- `google-calendar-disconnect`
- `google-calendar-create-event`
- `google-calendar-callback`

### Database update

Run the latest `supabase-schema.sql` again so these tables exist:

- `user_calendar_connections`
- `calendar_oauth_states`

## Email and text reminder setup

The app now supports one-click lead reminders through secure backend functions.

### What it does

- reps and admins can send reminder emails from a lead
- reps and admins can send reminder texts from a lead
- admins can edit the default email and SMS reminder templates in the Integrations tab
- sends are scoped to the lead the user is allowed to access

### Email provider setup

This build uses Resend for email.

Add these Supabase Edge Function secrets:

- `RESEND_API_KEY`
- `REMINDER_FROM_EMAIL`
- `REMINDER_REPLY_TO` (optional)

### SMS provider setup

This build uses Twilio for SMS.

Add these Supabase Edge Function secrets:

- `TWILIO_ACCOUNT_SID`
- `TWILIO_AUTH_TOKEN`
- `TWILIO_FROM_NUMBER`

### Functions to deploy

- `communication-status`
- `send-reminder`

## Recommended next improvements

- admin invite flow inside the app
- lead reassignment tools in bulk
- Kanban pipeline view for reps
- timeline/activity log for each lead
- reminder notifications for overdue follow-up
