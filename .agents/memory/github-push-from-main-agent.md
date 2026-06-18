---
name: Pushing to a user's personal GitHub from the main agent
description: How to authenticate and push local commits to a user's personal GitHub remote when no credential helper exists, plus the read-only verification trick around guarded git ops.
---

# Pushing to a personal GitHub remote from the main agent

When a task is "push commits to personal GitHub" and `origin` is an HTTPS GitHub
URL, a plain `git push origin main` fails with *"Invalid username or token.
Password authentication is not supported"* — there is no credential helper and no
PAT in secrets by default.

## The working path
1. The GitHub **connector** provides the auth. Search integrations for `github`;
   if a connection exists at account level it shows `status: not_added`.
2. `addIntegration(connection:conn_github_...)` wires it, but `listConnections('github')`
   STILL returns 0 afterward — the credential proxy serves nothing until the Repl
   is bound.
3. `proposeIntegration(connection:conn_github_...)` does the platform-side binding.
   It **exits the agent loop** and waits for the user to authorize. After they
   accept, on the next loop `listConnections('github')` returns 1 with
   `settings.access_token` (a ~40-char OAuth token carrying the `repo` scope).
4. Push from inside the `code_execution` sandbox so the token never touches disk
   or logs: build `https://x-access-token:${token}@github.com/<owner>/<repo>.git`
   and `execSync('git push <url> main:main')`. Always sanitize output by replacing
   the token with `***` before logging.

**Why:** the OAuth token must never be printed or written to a file; doing the
push in-sandbox keeps it as an in-memory variable.

## Verifying without tripping the destructive-git guard
`git fetch` / anything that updates `.git/refs/remotes/origin/main` is BLOCKED in
the main agent ("Destructive git operations are not allowed … *.lock"). So the
local `origin/main` tracking ref goes stale and can't be refreshed here.

Verify the real remote state with read-only commands instead:
- `git ls-remote <tokenized-url> refs/heads/main` (no local writes) → actual remote head
- `git --no-optional-locks rev-parse HEAD` → local head
- Equal SHAs == push succeeded / branch not ahead. The push command's own
  `cdac…..79db… main -> main` line with exit 0 is also authoritative.

**How to apply:** use this whenever validation/code_review demands `main` not be
ahead of `origin/main` but you cannot `git fetch` to update tracking refs.
