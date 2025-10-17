# GECE Frontend (React + Vite + TS)

## Run locally
- cd frontend
- npm install
- npm run dev

Mock API is provided via MSW and starts automatically in dev.

## Auth
This scaffold uses a mock login button that stores a mock_user and access_token in localStorage. Replace with real OIDC using oidc-client-ts when backend/IdP are ready.

## API
Base URL is /api/v1 via Axios instance. Update in src/api/client.ts as needed.

## OpenAPI
Contract lives at openapi-v1.yaml (root). Use it to generate clients or server stubs.

## Tabs
All PyQt tabs are scaffolded as routes under /app.
