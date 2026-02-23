> [!WARNING]
> This project is unmaintained and no longer actively developed.

# Excel AI Agent

A TypeScript-based AI agent for Excel operations using Claude LLM and Deno.

## Project Structure

```
src/
├── services/
│   ├── llm/
│   │   └── claude.ts          # Claude LLM service
│   └── agent/
│       └── excel-agent.ts     # Main AI agent orchestrator
├── tools/
│   ├── excel/
│   │   └── read-range.ts      # Excel read range tool
│   └── index.ts               # Tool registry
├── types/
│   └── index.ts               # TypeScript type definitions
└── utils/                     # Utility functions
tests/                         # Test files
```

## Setup

1. Copy `.env.example` to `.env` and add your Anthropic API key:

   ```
   ANTHROPIC_API_KEY=your_key_here
   ```

2. Run the agent:

   ```bash
   deno task start
   ```

## Features

- Claude LLM integration for natural language processing
- Tool system for Excel operations
- Mock `read_range()` tool with 3-second delay
- Extensible architecture for adding more tools
- TypeScript with strict mode enabled

## Database Migrations

The project uses SQLite with a custom migration system to manage schema changes.

### How Migrations Work

- Migrations run automatically on server startup
- Migration state is tracked in the `schema_version` table
- Only pending migrations (version > current version) are executed
- All migrations run in transactions (safe rollback on failure)
- Existing data is never affected

### Adding a New Migration

1. Open [src/db/migrations.ts](src/db/migrations.ts)
2. Add a new migration to the `migrations` array:

```typescript
{
  version: 2,  // Increment from last version
  name: "descriptive_migration_name",
  up: (db: DB) => {
    // Your SQL changes here
    db.execute(`
      CREATE TABLE new_table (
        id INTEGER PRIMARY KEY,
        data TEXT
      );
    `);
  }
}
```

1. Restart the server - the migration runs automatically

### Current Schema

**messages** - Conversation history

- `id`, `workbook_key`, `conversation_id`, `role`, `content`, `created_at`

**llm_keys_config** - LLM API configuration

- `id`, `provider`, `azure_openai_endpoint`, `azure_openai_key`, `added_date`,
  `modified_date`

**schema_version** - Migration tracking (auto-managed)

- `version`, `name`, `applied_at`

### Data Access Layers

- [src/db/dal.ts](src/db/dal.ts) - Message operations
- [src/db/llm-keys-dal.ts](src/db/llm-keys-dal.ts) - LLM keys configuration

## Available Tools

- `read_range`: Read data from Excel ranges (currently mocked)

## Development

- `deno task dev`: Run with file watching
- `deno task test`: Run tests

## local langfuse

- git clone <https://github.com/langfuse/langfuse.git>
- cd langfuse
- docker-compose up -d
- connect to <http://localhost:3000>
- signup for a new account + create project + API key
- add local stuff to .env:
  - LANGFUSE_BASE_URL=<http://localhost:3000>
  - LANGFUSE_SECRET_KEY=
  - LANGFUSE_PUBLIC_KEY=
