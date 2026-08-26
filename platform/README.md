# Oneiro Platform

Multi-tenant operations automation for pavement marking and construction subcontractors.

## Local Development

### Prerequisites
- Node.js 20+
- Python 3.12+
- Docker (for PostgreSQL + Redis)

### Setup

```bash
# Start infrastructure
docker compose up -d

# API server
cd api
cp .env.example .env     # Edit with your credentials
npm install
npm run db:push           # Push schema to local database
npm run db:seed           # Seed NYC DOT defaults
npm run dev               # Start dev server on :3001

# Worker (separate terminal)
cd worker
pip install -r requirements.txt
python main.py

# Frontend (separate terminal)
cd web
npm install
npm run dev               # Start Vite on :5173
```

### Architecture
- **api/** — TypeScript/Node Express API server
- **worker/** — Python PDF fill worker (consumes Redis jobs)
- **web/** — React 18 SPA (Vite)
- **migrations/** — SQL migration files
- **seed/** — Default data (NYC DOT categories, multipliers, classifications)

See `docs/mvp_design_plan.md` for full architecture documentation.
