#!/usr/bin/env bash
# Project .claude/skills/ is the source of truth (versioned in git).
# Copy to ~/.claude/skills/ so the skills also trigger from other pitch folders,
# and to the shared skills library so they travel with the rest of the collection.
set -e
cd "$(dirname "$0")"
LIB="D:/Personal/Skills_aug2026/skills-main/skills"
for s in brief-to-plan competitor-comms-audit category-creative-scan campaign-concept communication-strategy comment-analysis tg-profile; do
  rm -rf "$HOME/.claude/skills/$s"
  cp -r "$s" "$HOME/.claude/skills/"
  echo "synced $s -> ~/.claude/skills/"
  if [ -d "$LIB" ]; then
    rm -rf "$LIB/$s"
    cp -r "$s" "$LIB/"
    echo "synced $s -> $LIB/"
  fi
done
