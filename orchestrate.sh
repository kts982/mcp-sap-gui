#!/usr/bin/env bash
# orchestrate.sh — Helper for multi-CLI task delegation
# Place in your project root alongside CLAUDE.md
#
# Usage:
#   ./orchestrate.sh <codex|gemini> <task-name> [extra-files...]
#
# Prerequisites:
#   - Task spec must exist at .tasks/<task-name>.task.md
#   - Result is written to .tasks/<task-name>.result.md
#
# Examples:
#   ./orchestrate.sh codex generate-tests
#   ./orchestrate.sh gemini review-pr src/main.py src/utils.py
#   ./orchestrate.sh gemini analyze-logs logs/error.log

set -uo pipefail

ENGINE="${1:?Usage: orchestrate.sh <codex|gemini> <task-name> [files...]}"
TASK_NAME="${2:?Usage: orchestrate.sh <codex|gemini> <task-name> [files...]}"
shift 2
EXTRA_FILES=("$@")

TASK_DIR=".tasks"
TASK_FILE="${TASK_DIR}/${TASK_NAME}.task.md"
RESULT_FILE="${TASK_DIR}/${TASK_NAME}.result.md"

mkdir -p "$TASK_DIR"

if [[ ! -f "$TASK_FILE" ]]; then
  echo "ERROR: Task spec not found: $TASK_FILE"
  echo "Write the task spec first, then run this script."
  exit 1
fi

TASK_CONTENT=$(cat "$TASK_FILE")
echo "[$(date +%H:%M:%S)] Delegating '${TASK_NAME}' to ${ENGINE}..."

case "$ENGINE" in
  codex)
    # Use exec subcommand for non-interactive mode
    codex exec --full-auto \
      "$TASK_CONTENT" > "$RESULT_FILE" 2>&1
    ;;

  gemini)
    # Build file context: prepend file contents to the prompt
    FILE_CONTEXT=""
    for f in "${EXTRA_FILES[@]}"; do
      if [[ -f "$f" ]]; then
        FILE_CONTEXT+="--- File: ${f} ---"$'\n'
        FILE_CONTEXT+="$(cat "$f")"$'\n\n'
      else
        echo "WARNING: File not found, skipping: $f"
      fi
    done
    FULL_PROMPT="${FILE_CONTEXT}${TASK_CONTENT}"
    gemini -p "$FULL_PROMPT" --sandbox false > "$RESULT_FILE" 2>&1
    ;;

  *)
    echo "ERROR: Unknown engine '${ENGINE}'. Use 'codex' or 'gemini'."
    exit 1
    ;;
esac

EXIT_CODE=$?

if [[ $EXIT_CODE -eq 0 ]]; then
  LINES=$(wc -l < "$RESULT_FILE")
  echo "[$(date +%H:%M:%S)] Done. Result: ${RESULT_FILE} (${LINES} lines)"
else
  echo "[$(date +%H:%M:%S)] FAILED (exit ${EXIT_CODE}). Partial output: ${RESULT_FILE}"
fi

exit $EXIT_CODE
