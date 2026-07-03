// ====================================================
// Quick Read Hotmail - Backend Server
// Node.js + Express + Microsoft Graph API
// ====================================================

const express = require('express');
const cors = require('cors');
const fetch = require('node-fetch');
const path = require('path');

const app = express();
const PORT = process.env.PORT || 3000;

// Middleware
app.use(cors());
app.use(express.json({ limit: '2mb' }));

// Request Logger
app.use((req, res, next) => {
  console.log(`[${new Date().toISOString()}] ${req.method} ${req.url} - IP: ${req.ip}`);
  next();
});

app.use(express.static(path.join(__dirname, 'public')));

// ====================================================
// 1) Get new access_token using refresh_token
// ====================================================
async function getAccessToken(refreshToken, clientId) {
  const tokenUrl = 'https://login.microsoftonline.com/common/oauth2/v2.0/token';

  const params = new URLSearchParams();
  params.append('client_id', clientId);
  params.append('refresh_token', refreshToken);
  params.append('grant_type', 'refresh_token');
  params.append('scope', 'https://graph.microsoft.com/Mail.Read offline_access');

  const res = await fetch(tokenUrl, {
    method: 'POST',
    headers: { 'Content-Type': 'application/x-www-form-urlencoded' },
    body: params
  });

  const data = await res.json();
  if (!data.access_token) {
    throw new Error(data.error_description || 'Failed to get access token');
  }
  return data.access_token;
}

// ====================================================
// 2) Fetch inbox messages using Microsoft Graph API
// ====================================================
async function fetchInbox(accessToken, top = 20) {
  const url = `https://graph.microsoft.com/v1.0/me/mailFolders/inbox/messages?$top=${top}&$select=id,subject,from,bodyPreview,receivedDateTime,isRead,body`;

  const res = await fetch(url, {
    headers: {
      'Authorization': `Bearer ${accessToken}`,
      'Accept': 'application/json'
    }
  });

  const data = await res.json();
  if (data.error) {
    throw new Error(data.error.message || 'Failed to fetch emails');
  }
  return data.value || [];
}

// ====================================================
// API: Read emails for one or multiple accounts
// POST /api/read
// Body: { credentials: "email|password|refresh_token|client_id\n..." }
// ====================================================
app.post('/api/read', async (req, res) => {
  const { credentials } = req.body;

  if (!credentials || typeof credentials !== 'string') {
    return res.status(400).json({ error: 'Credentials are required' });
  }

  const lines = credentials.split('\n').map(l => l.trim()).filter(Boolean);
  const results = [];

  // Process all accounts in parallel for speed
  await Promise.all(lines.map(async (line) => {
    const parts = line.split('|').map(p => p.trim());

    if (parts.length < 4) {
      results.push({
        email: parts[0] || 'unknown',
        success: false,
        error: 'Invalid format. Expected: email|password|refresh_token|client_id',
        messages: []
      });
      return;
    }

    const [email, password, refreshToken, clientId] = parts;

    try {
      const accessToken = await getAccessToken(refreshToken, clientId);
      const messages = await fetchInbox(accessToken, 20);

      results.push({
        email,
        success: true,
        count: messages.length,
        messages: messages.map(m => ({
          id: m.id,
          subject: m.subject || '(No subject)',
          from: m.from?.emailAddress?.name || 'Unknown',
          fromEmail: m.from?.emailAddress?.address || '',
          preview: m.bodyPreview || '',
          date: m.receivedDateTime,
          isRead: m.isRead,
          body: m.body?.content || m.bodyPreview || '',
          bodyType: m.body?.contentType || 'text'
        }))
      });
    } catch (err) {
      results.push({
        email,
        success: false,
        error: err.message,
        messages: []
      });
    }
  }));

  res.json({
    accounts: results.length,
    totalMessages: results.reduce((sum, r) => sum + (r.messages?.length || 0), 0),
    results
  });
});

// Health check
app.get('/api/health', (req, res) => {
  res.json({ status: 'ok', service: 'hotmail-reader' });
});

// Serve frontend
app.get('/', (req, res) => {
  res.sendFile(path.join(__dirname, 'public', 'index.html'));
});

app.listen(PORT, '0.0.0.0', () => {
  console.log(`✅ Hotmail Reader running on http://0.0.0.0:${PORT}`);
});

// Process Error Listeners
process.on('uncaughtException', (err) => {
  console.error('❌ [UNCAUGHT EXCEPTION]', err);
});

process.on('unhandledRejection', (reason, promise) => {
  console.error('❌ [UNHANDLED REJECTION]', reason);
});
