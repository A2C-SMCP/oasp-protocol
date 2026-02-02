# CLAUDE.md

This file provides guidance to Claude Code (claude.ai/code) when working with code in this repository.

## Project Overview

OASP (Office AddIn Socket Protocol) is a Socket.IO-based communication protocol specification for enabling AI Agents to control Microsoft Office documents through Office AddIns. This is a **documentation-only repository** - no application code, only protocol specifications.

## Commands

```bash
# Install dependencies (using uv)
uv sync

# Serve documentation locally with hot reload
inv docs.serve

# Build static documentation
inv docs.build

# Deploy documentation with versioning
inv docs.deploy --version=latest
```

## Architecture

### Three-Layer System

```
AI Agent ─[MCP/API]─► Server ─[Socket.IO]─► Office AddIn ─[Office.js]─► Microsoft Office
```

1. **AI Agent Layer**: LLM-based agents controlling Office via MCP/API
2. **Server Layer**: Python backend (Office4AI) with Socket.IO server
3. **AddIn Layer**: Office plugins using Socket.IO client and Office.js API

### Namespace Isolation

| Namespace | Application | Status |
|-----------|-------------|--------|
| `/word` | Microsoft Word | ✅ Stable |
| `/ppt` | PowerPoint | 📋 Draft |
| `/excel` | Excel | 📋 Draft |

## Protocol Conventions

### Naming

- **JSON fields**: camelCase (`requestId`, `documentUri`)
- **Event names**: kebab-with-colon (`word:get:selection`)
- **Error codes**: SCREAMING_SNAKE_CASE (`SELECTION_EMPTY`)

### Event Name Format

`{namespace}:{action}:{target}` (e.g., `word:insert:text`)

### Request/Response Pattern

All requests include `requestId` (UUID v4), `documentUri`, and `timestamp` (Unix ms UTC). Responses include `success` boolean and either `data` or `error` object.

### Error Code Ranges

- 1xxx: Generic errors
- 2xxx: Connection/Auth errors
- 3xxx: Document/Operation errors
- 4xxx: Validation errors

## Documentation Structure

- `docs/specification/`: Protocol specifications (architecture, connection, events, data structures)
- `docs/appendix/`: Glossary and changelog
- `scripts/docs/tasks.py`: Invoke task definitions

When extending the protocol, update the relevant specification file and add entries to `docs/appendix/changelog.md` with appropriate status markers (✅ Stable, 📋 Draft, ⚠️ Deprecated).
