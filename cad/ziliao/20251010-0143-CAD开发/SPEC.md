Title: <Concise task name>
Owner: <name>
Created: <YYYY-MM-DD>
Status: Draft | In-Progress | Final

1. Problem & Goals
- Problem: <what hurts / why now>
- Primary goal(s): <outcomes, not activities>
- Non-goals: <explicitly out of scope>

2. Scope
- In-scope: <systems, repos, files, endpoints>
- Out-of-scope: <areas excluded>
- Constraints: <time, compatibility, security, approvals>

3. Requirements
- Functional: <bullet list of behaviors>
- Non-functional: <perf, reliability, usability>
- Acceptance criteria: <observable checks with examples>

4. Interfaces
- CLI/API contracts: <commands, routes, payloads>
- Data model: <schemas, fields, types>
- External tools/MCP: <tool names, capabilities, limits>

5. Plan
- Milestones:
  1) <milestone name>: definition of done
  2) <milestone name>: definition of done
- Work breakdown (maps to update_plan steps):
  - <step 1>
  - <step 2>

6. Validation
- Test strategy: <unit/e2e/manual>
- Test cases: <cases with input/output/paths>
- Rollback/Backout: <how to revert>

7. Deliverables
- Artifacts: <files/paths>
- Docs: <README updates, runbook>
- Demo: <how to run>

8. Risks & Mitigations
- <risk> -> <mitigation>

9. Directory Layout
- root/  (task root)
  - src/
  - data/
  - artifacts/
  - logs/
  - SPEC.md  (this file)
  - AGENTS.md

10. Runbook
- Bootstrap: <commands to install deps>
- Dev loop: <commands>
- Validation: <commands>

Appendix
- References/links
- Glossary
