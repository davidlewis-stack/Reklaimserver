const express = require('express');
const fs = require('fs');
const path = require('path');
const { execSync } = require('child_process');
const crypto = require('crypto');

const app = express();
app.use(express.json({ limit: '2mb' }));

app.use((req, res, next) => {
  res.header('Access-Control-Allow-Origin', '*');
  res.header('Access-Control-Allow-Methods', 'POST, OPTIONS');
  res.header('Access-Control-Allow-Headers', 'Content-Type');
  if (req.method === 'OPTIONS') return res.sendStatus(200);
  next();
});

app.post('/generate', async (req, res) => {
  const { config } = req.body;
  if (!config) return res.status(400).json({ error: 'No config provided' });

  const id = crypto.randomBytes(8).toString('hex');
  const configPath = `/tmp/deck_config_${id}.json`;
  const outPath = `/tmp/Reklaim_${id}.pptx`;

  try {
    const cfg = { ...config, outputPath: outPath };
    fs.writeFileSync(configPath, JSON.stringify(cfg, null, 2));

    execSync(
      `node /home/claude/reklaim-app/reklaim_deck_template.js ${configPath}`,
      { cwd: '/home/claude/reklaim-app', timeout: 60000 }
    );

    const fileData = fs.readFileSync(outPath);
    const filename = `Reklaim_${(config.brand || 'Deck').replace(/\s+/g, '_')}_${(config.shortVertical || 'Pitch').replace(/\s+/g, '_')}.pptx`;

    res.setHeader('Content-Type', 'application/vnd.openxmlformats-officedocument.presentationml.presentation');
    res.setHeader('Content-Disposition', `attachment; filename="${filename}"`);
    res.send(fileData);
  } catch (err) {
    res.status(500).json({ error: err.message });
  } finally {
    try { fs.unlinkSync(configPath); } catch {}
    try { fs.unlinkSync(outPath); } catch {}
  }
});

const PORT = 3456;
app.listen(PORT, () => console.log(`Reklaim deck server running on ${PORT}`));
