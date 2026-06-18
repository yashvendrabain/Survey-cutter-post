#!/bin/bash
set -e
pnpm install --frozen-lockfile
pnpm --filter db push
# Realign .pythonlibs to uv.lock so the python-lock validation stays green
# after a lock bump (no-op when already in sync). See post_merge_setup skill.
.pythonlibs/bin/python artifacts/survey-insight-engine/scripts/realign_python_lock.py
