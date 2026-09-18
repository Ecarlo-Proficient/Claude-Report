# CLAUDE.md patch — add AI Brain_Vault entry

**Date:** 2026-05-05
**Reason:** Document the AI Brain_Vault location so future sessions know it's a writable workspace distinct from the read-only Obsidian Main Vault.

---

## Apply this change

In your Cowork global CLAUDE.md, under the **Tools & Systems** section, between the Obsidian vault entry and the Project outputs entry, add this bullet:

```markdown
- **AI Brain_Vault** (`~/Documents/Claude/AI Brain_Vault`): Writable working space for AI-collaborative knowledge — index, running log, per-task notes. Drafts intended for the read-only Obsidian Main Vault go to `AI Brain_Vault/drafts/vault-fills/` mirroring vault structure. Has its own `.obsidian/` so Obsidian can open it as a vault. NOT the Main Vault — Main Vault remains read-only and untouched.
```

The Obsidian Main Vault entry stays exactly as it is — that vault is still read-only and at the same path.

---

## Where the global CLAUDE.md lives

Cowork loads the user-global CLAUDE.md from your persistent cowork session memory directory. Easiest way to edit it: open it via the file path shown in the system prompt of any Cowork session, or use whatever local mirror you maintain.

If you maintain a local source-of-truth copy (e.g. in your Documents/Claude folder), edit that and let your usual sync process do its thing. If you edit the active in-session copy directly, the change persists across future sessions.

---

## Verification after applying

In the next new Cowork session, you should see the AI Brain_Vault entry in the Tools & Systems block of the loaded CLAUDE.md context. If you don't, the edit didn't land in the persistent file.
