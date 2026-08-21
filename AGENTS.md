# Repository Instructions for AI Agents

Before changing this repository, read `AI_START_HERE.md` and every document it lists.

- `docs/铁律.md` is the single source of truth for rules. Do not restate its rules in other documents — link to it instead.

- Keep the repository usable by assistants with no prior chat context.
- Do not commit customer PII, tickets, passports, order screenshots, or private chat logs.
- Do not add customer-communication scripts. Talking to customers is the operator's job, not the assistant's.
- Keep local absolute paths in `config/local-paths.yaml`, which is ignored by Git.
- Update catalog entries when adding or changing a city template summary.
- Do not create a separate cross-city template for every possible city pair.
- Preserve the distinction between confirmed requirements, optional suggestions, and reusable template content.
- Original DOCX templates are changed only after human approval.

