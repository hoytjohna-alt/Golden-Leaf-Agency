# Golden Leaf Agency HQ

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

## Current admin capabilities

- full agency dashboard
- cross-rep opportunity visibility
- rep assignment on leads
- coaching note management
- global assumptions and carrier commission edits
- profile role and active-status management inside the app
- password reset request from the login screen

## Recommended next improvements

- admin invite flow inside the app
- lead reassignment tools in bulk
- Kanban pipeline view for reps
- timeline/activity log for each lead
- reminder notifications for overdue follow-up
