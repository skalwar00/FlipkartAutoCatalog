# Deployment Guide (Vercel + Railway)

This project is set up for:

- Frontend on Vercel
- Backend on Railway

## 1) Deploy Backend on Railway

Use the repository root as the Railway service root (important because backend uses workspace packages).

### Build Command

```bash
pip install -r backend/requirements.txt && pnpm install --frozen-lockfile && pnpm --filter @workspace/api-server build
```

### Start Command

```bash
PORT=$PORT pnpm --filter @workspace/api-server start
```

### Required Environment Variables (Railway)

- `PORT` (Railway usually injects this automatically)
- `NODE_ENV=production`
- `FRONTEND_ORIGIN=https://<your-vercel-domain>`
- `DATABASE_URL=<your-db-connection-string>` (if required by your runtime paths)
- `PYTHON_PATH=python3` (optional, only if default python path differs)

After deploy, copy your Railway backend URL (example: `https://your-api.up.railway.app`).

## 2) Deploy Frontend on Vercel

Deploy from this project, with:

- Framework preset: `Vite`
- Root Directory: `frontend`
- Build command: `pnpm build`
- Output directory: `dist/public`

### Required Environment Variables (Vercel)

- `PORT=5173` (or any fixed port your config expects during build/dev flows)
- `BASE_PATH=/`
- `VITE_API_BASE_URL=https://<your-railway-backend-domain>`

## 3) DNS/CORS Checklist

- Ensure Railway `FRONTEND_ORIGIN` matches your deployed Vercel domain exactly.
- If you use a custom frontend domain, add that domain to `FRONTEND_ORIGIN`.
- If multiple frontend origins are needed, set comma-separated values.

## 4) Smoke Test

1. Open Vercel app.
2. Upload sample files in `CatalogMapper`.
3. Click generate/download.
4. Confirm backend responds at `https://<railway>/api/fill-xls` and file downloads.
