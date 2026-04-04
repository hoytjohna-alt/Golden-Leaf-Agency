# Golden Leaf Agency HQ

This version is set up for the workflow you described:

- each rep or producer gets their own login
- each rep maintains only their own lead pipeline
- every rep's activity feeds the shared agency dashboard
- the owner or admin gets the full master view across the business

## Architecture

- Frontend: static HTML/CSS/JavaScript
- Auth and database: Supabase
- Hosting: GitHub Pages or Render static hosting

That combination works well because the app can stay a simple static site while still getting:

- secure login
- row-level access control
- shared cloud data
- admin-only global settings

## Files

- `index.html` - app shell
- `styles.css` - layout and visual styling
- `app.js` - authenticated app logic and dashboard calculations
- `supabase-schema.sql` - database tables, trigger, and row-level security policies
- `supabase-config.js` - local config file for Supabase keys
- `supabase-config.example.js` - example config
- `Golden Leaf Agency Workbook (w_Comp).xlsx` - original workbook source

## Data model

The schema creates these core tables:

- `profiles`
  - one row per authenticated user
  - role is `admin` or `rep`
- `opportunities`
  - shared lead pipeline table
  - every lead is assigned to a specific user
- `coaching_notes`
  - weekly coaching notes by rep
- `app_settings`
  - agency-wide assumptions, statuses, products, lead sources, and carrier commission table

## Security model

Row-level security is included in `supabase-schema.sql`.

- reps can only `select`, `insert`, `update`, and `delete` opportunities assigned to themselves
- admins can access all opportunities
- reps can only see their own profile
- admins can see all profiles
- all authenticated users can read app settings
- only admins can change app settings
- coaching notes are readable by the rep they belong to and fully manageable by admins

## Setup

### 1. Create the Supabase project

1. Create a new Supabase project.
2. Open the SQL Editor.
3. Run the full contents of `supabase-schema.sql`.

### 2. Add your Supabase keys

Edit `supabase-config.js` and paste:

- `supabaseUrl`
- `supabaseAnonKey`

You can also copy from `supabase-config.example.js` if you want a fresh template.

### 3. Create user accounts

Create users in Supabase Auth for:

- agency owner/admin
- each rep/producer

The SQL trigger automatically creates their `profiles` row when a user is created.

After that, update the admin user's role:

```sql
update public.profiles
set role = 'admin'
where email = 'owner@youragency.com';
```

All other users can stay as `rep`.

Once users exist, the app now lets the admin:

- change a user's role between `rep` and `admin`
- mark a user active or inactive
- update display names

Users can also request their own password reset email from the login screen.

### 4. Deploy

#### GitHub Pages

1. Push this folder to a GitHub repository.
2. In GitHub, open `Settings` -> `Pages`.
3. Choose `Deploy from a branch`.
4. Select `main` and `/ (root)`.
5. Save.

#### Render

1. Create a new Static Site.
2. Connect the GitHub repo.
3. Leave the build command blank.
4. Set the publish directory to `.`
5. Deploy.

## How it behaves

### Rep view

- signs into a personal account
- sees only their own leads
- can create, edit, and manage their pipeline throughout the lifecycle
- can see their own dashboard and scorecard
- can view coaching notes that apply to them

### Admin view

- sees the full agency dashboard
- sees every producer's pipeline
- can assign leads to reps
- can review rep scorecards and lead source ROI
- can update coaching notes
- can change agency-wide assumptions and carrier commission settings

## Local preview

Because this app uses Supabase in the browser, the cleanest local preview is to serve the folder:

```bash
python3 -m http.server 8000
```

Then open [http://localhost:8000](http://localhost:8000).

## Next upgrade

This repo is now positioned for a true shared online app. The strongest next step would be:

- branded login screen
- admin user management inside the app
- activity timeline on each lead
- notifications for overdue follow-ups
- file attachments and call notes

## GitHub publish checklist

This folder is now prepared for publishing:

- `.gitignore` keeps `supabase-config.js` out of source control
- `render.yaml` gives Render a basic static-site config

To publish:

```bash
git init
git add .
git commit -m "Initial Golden Leaf Agency HQ app"
git branch -M main
git remote add origin YOUR_GITHUB_REPO_URL
git push -u origin main
```

If you use GitHub Pages, do not commit your real `supabase-config.js`. Keep that file local and set the production copy before deployment, or replace it with your hosted keys intentionally.
