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

## Phase 3 — The weekly coaching loop (your machine)

This keeps the bank growing on its own, aimed at her weak spots. Design (built to be safe with
teaching content — nothing wrong gets taught):

1. **`weekly-coach.ps1`** (in this folder) runs on a schedule. It **pulls her state**, works out
   which subjects and levels she's weakest in (verbatim from her ability scores), and writes a
   **coaching brief** (`coach-brief.json`) — e.g. "needs 4 new Chemistry L2, 3 Vet L3".
2. **I draft the questions** from that brief (the reasoning/content step) and **verify each answer**,
   then post them to the **`Pending`** tab — never straight to live.
3. **Quick review gate:** you (or I, in a session) glance at `Pending`, and good rows get copied
   into `Questions` with `active = TRUE`. That's the only step that puts a question in front of her.

To schedule the compute step, once the endpoint is set:
```
schtasks /Create /TN "Paige Road to Vet Coach" /TR "powershell -ExecutionPolicy Bypass -File C:\Users\BrendonGraham\paige-vet-pathway\weekly-coach.ps1" /SC WEEKLY /D SUN /ST 07:00
```
(Mirrors the feed-snapshot task pattern.)

**Decision still open:** whether the drafting step (2) runs through the AI Bridge automatically each
week or I do it when you nudge me. Both keep the review gate. Tell me which and I'll wire it.

---

## Where her memory lives
- **On her phone** (localStorage) — always, and it works with no internet.
- **In the Sheet's `State` tab** — a backup copy, once Phase 2 is set up. Newest copy wins, so she
  can move to a new phone and pull her progress back.
- It is **per person, keyed `paige`**. Nothing here is sensitive (subject scores and question ids only).
