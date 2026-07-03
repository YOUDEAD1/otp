# STOCK LARA OTP

Fast email reader for Hotmail/Outlook accounts using Microsoft Graph API and OAuth2 refresh tokens. Supports English and Arabic interfaces.

## Features

- Read emails from Hotmail/Outlook/Live accounts instantly
- Bilingual UI (English / Arabic) — switchable from the top-right
- Multiple accounts support (parallel fetching for speed)
- Format: `email|password|refresh_token|client_id`
- Shows sender, subject, date, preview, full body
- Auto-refresh every 10 seconds
- Live stats: accounts, messages, unread count
- Renders HTML emails correctly

## Local Setup

```bash
npm install
npm start
```

Then open `http://localhost:3000`

## Deploy to VPS (PM2)

```bash
sudo apt update && sudo apt install nodejs npm -y
sudo npm install -g pm2

cd /var/www/stock-lara-otp
npm install
pm2 start server.js --name stock-lara
pm2 startup
pm2 save
```

## Deploy with Docker

```bash
docker build -t stock-lara-otp .
docker run -d -p 3000:3000 --name stock-lara stock-lara-otp
```

## Nginx Reverse Proxy

```nginx
server {
    listen 80;
    server_name your-domain.com;
    location / {
        proxy_pass http://localhost:3000;
        proxy_http_version 1.1;
        proxy_set_header Host $host;
        proxy_set_header X-Real-IP $remote_addr;
    }
}
```

Add free HTTPS:
```bash
sudo certbot --nginx -d your-domain.com
```

## How it works

1. Frontend sends credentials to `/api/read`
2. Backend exchanges `refresh_token` for `access_token` from Microsoft
3. Backend calls Microsoft Graph API to fetch inbox
4. Returns JSON to the frontend

All accounts processed in parallel using `Promise.all()` — 10 accounts fetch in roughly the same time as 1.

## Endpoints

- `GET /` — Web UI
- `POST /api/read` — Body: `{ credentials: "email|password|refresh_token|client_id\n..." }`
- `GET /api/health` — Health check

## Security Notes

- No credentials are stored anywhere
- All requests go directly between your server and Microsoft
- Use HTTPS in production (free via Let's Encrypt)
- Consider adding rate limiting for public deployments

**STOCK LARA OTP** · Built for speed
