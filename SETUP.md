# Road to Vet — sync & coaching setup

The app works fully offline with nothing set up. This adds the cloud loop so her progress
is backed up, I can see how she's going, and her questions get topped up over time.

There are three phases. Phase 1 (the app) is done. Phases 2 and 3 need you, because they
touch your Google account and your machine.

---

## Phase 2 — Stand up the Sheet + sync (about 10 minutes, one-off)

1. **Create a Google Sheet** — call it `Paige Road to Vet`.
2. Make three tabs (bottom-left), named exactly:
   - `Questions` — first row: `id | subj | lvl | q | o1 | o2 | o3 | o4 | answer | ex | active`
   - `State` — first row: `id | updatedAt | json`
   - `Pending` — same header as `Questions`
   (If you skip this, the script creates them on first run — but making them yourself is clearer.)
3. **Extensions → Apps Script.** Delete the sample code, paste in `Code.gs` from this folder, Save.
4. **Deploy → New deployment → type "Web app".**
   - Execute as: **Me**
   - Who has access: **Anyone**  ← must be "Anyone", not "Anyone with a Google account" (anonymous fetch)
   - Deploy, authorise when prompted, and **copy the `/exec` URL**.
5. **Open the app → gear icon (top right) → paste the URL → Save.** It flips to "Syncing to the cloud".
   Do this on Paige's phone. Tell her to open it once on wifi so the first sync runs.

That's the backup + my visibility done. From here, her progress lands in the `State` tab
(one row, `id = paige`, a JSON blob with her ability per subject, streak, history and the
questions she's missed).

### Adding questions by hand (anytime, no redeploy)
Add a row to the `Questions` tab, set `active` to `TRUE`, and it shows up in her app on next open.
- `subj` = one of `bio chem phys vet math`
- `lvl` = 1, 2 or 3
- `answer` = 1–4 (which option is right)
- Use a new, unique `id` to add; reuse an existing `id` to overwrite that question.

---

## Phase 3 — The autonomous weekly loop (your machine)

**Mode: fully automatic, no approval step** (Brendon's call, 20 Aug 2026). The weekly job drafts
new questions for her weak subjects and publishes them straight to the live `Questions` tab.
The one safeguard kept is an **automated self-check** — each new question's answer is
independently re-derived by the model, and any that don't match are dropped. This is not an
approval gate (nothing waits on a human); it just stops a wrong answer key ever reaching her.

**`weekly-questions.ps1`** does the whole thing unattended:
1. Pulls her synced progress, finds her 3 weakest subjects and their working level.
2. Calls the Claude API to draft `perSubject` questions each, avoiding duplicates of what she has.
3. Re-checks every answer with a second, independent API call; keeps only the ones that verify.
4. Publishes the survivors to the live `Questions` tab (active = TRUE) — she sees them next open.
5. Logs what it did to `coach-log.md`.

(`weekly-coach.ps1` is the lighter compute-only version — brief to a file, no publishing. Kept as
a dry-run / diagnostic. `weekly-questions.ps1` is the live one.)

### Switch it on (3 steps)
1. **Redeploy `Code.gs`** — it now has a `publish` handler for the auto-loop. Paste the updated
   `Code.gs` over the old one in the Apps Script editor, then **Deploy → Manage deployments →
   edit → New version**. (The `/exec` URL stays the same.)
2. **Create `coach-config.json`** in this folder (copy `coach-config.example.json`) and fill in:
   - `endpoint` — your `/exec` URL
   - `anthropicKey` — your Claude API key
   (This file is git-ignored, so the key never leaves your machine.)
3. **Schedule it** (weekly, Sunday 7am):
```
schtasks /Create /TN "Paige Road to Vet Coach" /TR "powershell -ExecutionPolicy Bypass -File C:\Users\BrendonGraham\paige-vet-pathway\weekly-questions.ps1" /SC WEEKLY /D SUN /ST 07:00
```
Run it once by hand first (`powershell -ExecutionPolicy Bypass -File weekly-questions.ps1`) to
watch the first batch land, then leave it to the schedule.

---

## Where her memory lives
- **On her phone** (localStorage) — always, and it works with no internet.
- **In the Sheet's `State` tab** — a backup copy, once Phase 2 is set up. Newest copy wins, so she
  can move to a new phone and pull her progress back.
- It is **per person, keyed `paige`**. Nothing here is sensitive (subject scores and question ids only).
