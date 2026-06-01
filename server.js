const express = require('express');
const fs = require('fs');
const { execSync } = require('child_process');
const crypto = require('crypto');
const path = require('path');

const app = express();
app.use(express.json({ limit: '2mb' }));

app.use((req, res, next) => {
  res.header('Access-Control-Allow-Origin', '*');
  res.header('Access-Control-Allow-Methods', 'POST, GET, OPTIONS');
  res.header('Access-Control-Allow-Headers', 'Content-Type');
  if (req.method === 'OPTIONS') return res.sendStatus(200);
  next();
});

app.get('/', (req, res) => res.json({ status: 'ok', service: 'reklaim-deck-server' }));

app.get('/app', (req, res) => {
  res.sendFile(path.join(__dirname, 'index.html'));
});

app.post('/claude', async (req, res) => {
  const apiKey = process.env.ANTHROPIC_API_KEY;
  if (!apiKey) return res.status(500).json({ error: 'API key not configured on server' });
  try {
    const response = await fetch('https://api.anthropic.com/v1/messages', {
      method: 'POST',
      headers: { 'Content-Type': 'application/json', 'x-api-key': apiKey, 'anthropic-version': '2023-06-01' },
      body: JSON.stringify(req.body),
    });
    const data = await response.json();
    res.json(data);
  } catch (err) {
    res.status(500).json({ error: err.message });
  }
});

// Ensure a value is an array — converts objects with numeric keys to arrays
function ensureArray(val) {
  if (Array.isArray(val)) return val;
  if (val && typeof val === 'object') return Object.values(val);
  return [];
}

// Sanitize config to guarantee all array fields are actual arrays
function sanitizeConfig(cfg) {
  if (!cfg) return cfg;
  if (cfg.userJourney) cfg.userJourney = ensureArray(cfg.userJourney);
  if (cfg.catalogue) {
    cfg.catalogue.left = ensureArray(cfg.catalogue.left);
    cfg.catalogue.right = ensureArray(cfg.catalogue.right);
  }
  if (cfg.verticals?.cards) cfg.verticals.cards = ensureArray(cfg.verticals.cards);
  if (cfg.tier1?.audiences) cfg.tier1.audiences = ensureArray(cfg.tier1.audiences);
  if (cfg.tier2?.stats) cfg.tier2.stats = ensureArray(cfg.tier2.stats);
  if (cfg.tier2?.blocks) cfg.tier2.blocks = ensureArray(cfg.tier2.blocks);
  if (cfg.advantages?.items) cfg.advantages.items = ensureArray(cfg.advantages.items);
  if (cfg.solution?.steps) cfg.solution.steps = ensureArray(cfg.solution.steps);
  if (cfg.activation?.channels) cfg.activation.channels = ensureArray(cfg.activation.channels);
  return cfg;
}

app.post('/generate', async (req, res) => {
  let { config } = req.body;
  if (!config) return res.status(400).json({ error: 'No config provided' });
  config = sanitizeConfig(config);

  const id = crypto.randomBytes(8).toString('hex');
  const configPath = `/tmp/deck_config_${id}.json`;
  const outPath = `/tmp/Reklaim_${id}.pptx`;
  const templatePath = path.join(__dirname, 'reklaim_deck_template.js');

  try {
    // Use AI-generated contact only if it looks real (has reklaimyours.com domain), otherwise default to Neil
    const aiContact = config.contact || {};
    const emailLooksReal = aiContact.email && aiContact.email.includes("reklaimyours.com");
    const cfg = {
      ...config,
      outputPath: outPath,
      contact: {
        name: emailLooksReal ? aiContact.name : "Reklaim Sales Team",
        email: emailLooksReal ? aiContact.email : "neil@reklaimyours.com",
        web: "reklaimyours.com"
      }
    };
    fs.writeFileSync(configPath, JSON.stringify(cfg, null, 2));
    execSync(`node ${templatePath} ${configPath}`, { cwd: __dirname, timeout: 60000 });

    const fileData = fs.readFileSync(outPath);
    const filename = `Reklaim_${(config.brand || 'Deck').replace(/\s+/g, '_')}_${(config.shortVertical || 'Pitch').replace(/\s+/g, '_')}.pptx`;

    res.setHeader('Content-Type', 'application/vnd.openxmlformats-officedocument.presentationml.presentation');
    res.setHeader('Content-Disposition', `attachment; filename="${filename}"`);
    res.send(fileData);
  } catch (err) {
    console.error('Generation error:', err.message);
    res.status(500).json({ error: err.message });
  } finally {
    try { fs.unlinkSync(configPath); } catch {}
    try { fs.unlinkSync(outPath); } catch {}
  }
});

const PORT = process.env.PORT || 3456;
app.listen(PORT, () => console.log(`Reklaim deck server running on port ${PORT}`));
