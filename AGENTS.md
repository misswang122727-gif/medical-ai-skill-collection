# AGENTS.md

## Cursor Cloud specific instructions

### Repository overview

This is a **content-only repository** — it contains no runnable application code, no build system, and no automated tests. The repository hosts a curated collection of medical AI evaluation skills (prompt engineering assets) for the ByteDance Xpert benchmark platform.

The primary artifact is `skill/Health_Program_Expert_v2.skill`, a ZIP archive containing:
- **Prompt templates** (`prompts/`): system prompt, user prompt template, few-shot examples, fallback prompt
- **Reference knowledge bases** (`references/`): lab ranges, drug info, risk assessment, nutrition data, health plans, boundary responses, etc.
- **Metadata** (`SKILL.md`): skill name, version, and description

### Development workflow

- There are **no dependencies to install**, no package managers, and no build steps.
- There are **no lint checks, automated tests, or CI pipelines**.
- To inspect the `.skill` file contents: `unzip -l skill/Health_Program_Expert_v2.skill`
- To extract for editing: `unzip skill/Health_Program_Expert_v2.skill -d /tmp/skill-extract`
- To repackage after editing: `cd /tmp/skill-extract/health-expert-skill && zip -r /workspace/skill/Health_Program_Expert_v2.skill .`
- Content is primarily in Chinese (Simplified).

### Validation

A quick validation can be done with Python (available in the VM):
```bash
python3 -c "import zipfile; zf=zipfile.ZipFile('skill/Health_Program_Expert_v2.skill'); print('OK' if zf.testzip() is None else 'CORRUPT')"
```

The JSON file (`prompts/few_shot_examples.json`) can be validated with:
```bash
python3 -c "import zipfile,json; zf=zipfile.ZipFile('skill/Health_Program_Expert_v2.skill'); [print(json.loads(zf.read(n))) for n in zf.namelist() if n.endswith('.json')]"
```
