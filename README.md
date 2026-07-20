# Klaus — shared household organizer

The house, between you. Phase one: full app + AI drafting + photos + shared Supabase store.
(Outlook wiring is phase two, pocket push notifications phase three.)

## Deploy
This repo IS the deployed site — no build step on Netlify.
- Netlify settings: build command **(leave empty)** · publish directory **/** (root)
- Every push to main redeploys automatically.

## Files
- `index.html`, `app.js`, `manifest.json`, `icon-*.png` — the site itself
- `src/app.jsx` — the single-file source. **Edit this**, then rebuild `app.js`:
  `npx esbuild src/app.jsx --bundle --minify --outfile=app.js --loader:.jsx=jsx --define:process.env.NODE_ENV='"production"'`
- `supabase-setup.sql` — run once in the Supabase SQL Editor (see below)

## Supabase (one-time)
1. supabase.com → New project → name `klaus`, set a DB password, region West EU
2. SQL Editor → paste `supabase-setup.sql` → Run
3. Storage → New bucket → name `photos` → tick Public
4. Project Settings → API → copy Project URL + anon public key
5. In the app: gear → Shared store → paste both → Connect & sync (once per phone)

## Per-device keys (in the app, under the gear)
- Claude API key → enables photo drafting, note reading, after-photo verification
- Supabase URL + anon key → enables sync between phones
