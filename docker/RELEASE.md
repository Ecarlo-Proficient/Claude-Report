# Releasing to the office server (Synology)

**Who does what.** A release is THE process of the owner and Claude: we write the code, we check the Synology is
ready for it, we release it. The developer's part is the Synology side only - he carries out the instructions we
send (create a folder, add a setting, re-paste the update script, run the update). The checks are never his job;
a release that reaches him unchecked is our mistake (10/16/2026: a new folder for the SharePoint approval report
was missing, and the update took the server down).

Code reaches the Synology in one order: **dev → release PR into main → the update task on the Synology**. What the
server needs that code cannot bring - folders, Cloud Sync copies, `server.env` settings, the pasted update script -
is read from `compose.yml` and `server.env.example` themselves by one rule, `shared/server_release.py`, used in three
places:

| Where | Who | What it does |
|---|---|---|
| Ledger: gear → Office server → **Waiting for the server** | the owner | Synology ready / NOT ready for what is on dev, item by item, + the developer's instructions (Copy) |
| `docker/release_check.py` | Claude, before every release | the same check in a terminal (exit 1 = not ready) + the developer's instructions |
| `docker/update_server.sh` on the Synology | runs by itself | the last guard: stops before any change if anything is missing - the server keeps running |

## 1. Before the release PR is merged (the owner + Claude)

1. Run the check - ledger gear → **Refresh view**, or

   ```bash
   cd "<repo>" && bash -c '. python-env/python.sh && "$ACB_PY" docker/release_check.py'
   ```

2. **NOT ready** → send the developer the instructions it wrote (gear: *Instructions for the developer* → Copy;
   terminal: the block under *for the developer*). They are ready to forward as-is.
3. Run the check again once he is done. Merge only when it says **ready**. (*can't check from this Mac* = the
   encrypted keys folder; it already exists unless the instructions say otherwise.)
4. Merge the release PR (one approving review + green CI).

## 2. Putting it on the server (the developer, from the instructions)

5. If the update script changed, re-paste it in DSM (the instructions say so). Then run the update task
   (Task Scheduler → the update task → **Run**), or let the nightly 03:00 run take it.
6. The end of `update_server.log` (`/volume1/docker/automation/`; from a Mac `/Volumes/docker/automation/`):

   | Last lines | Meaning | Next |
   |---|---|---|
   | `office-server is running` · `done` | released | step 7 |
   | `STOPPED before any change` + `MISSING ...` | nothing pulled, the server untouched | set up what it names, **Run** again |
   | `THE SERVER IS DOWN` | the container did not start | fix the error just above it, **Run** again now |
   | `the update did NOT go in` | Docker failed (e.g. a build error) but the old container still runs | fix the error above it, **Run** again |
   | `WARNING - the server runs without` | running, a setting missing | add the setting, **Run** again |

## 3. After (the owner + Claude)

7. Ledger gear → Office server: **Running**, *The server runs* = the newest commit on main, nothing under
   *Next update*. The first status after a restart can take up to 3 minutes.

## If the server is down

- Gear shows **Down** (no word for an hour) or the log says `THE SERVER IS DOWN`.
- What the log names gets fixed, then the update task is **Run**: with no new code it skips the rebuild and starts
  the container (it still checks the folders first).
- The trackers and Notion are not written while it is down; the Mac's sync-all stays standing down while
  `writer.json` says server. For a long outage, `writer.json` goes to `mac` (`docker/README.md` step 8) and sync-all
  runs on the Mac until the server is back.
