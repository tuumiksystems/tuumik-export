/* Copyright (C) 2017-2025 Tuumik Systems OÜ */

import express from 'express';
import { xlsx1 } from './xlsx1/xlsx1.js';
import { pdf1 } from './pdf1/pdf1.js';

const app = express();
const PORT = 3000;

app.use(express.json());

app.post('/xlsx1', async (req, res) => {
  try {
    await xlsx1(req, res);
  } catch {
    res.status(500).send('Error generating XLSX file');
  }
});

app.post('/pdf1', async (req, res) => {
  try {
    await pdf1(req, res);
  } catch {
    res.status(500).send('Error generating PDF file');
  }
});

app.listen(PORT, () => {
  console.log(`Server running on http://localhost:${PORT}`);
});
