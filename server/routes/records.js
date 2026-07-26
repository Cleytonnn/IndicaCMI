import express from 'express';
import { v4 as uuidv4 } from 'uuid';
import { getAll, getById, create, update, remove } from '../db.js';

const router = express.Router();

router.get('/', async (req, res) => {
  try {
    const rows = await getAll();
    res.json(rows);
  } catch (err) {
    res.status(500).json({ error: err.message });
  }
});

router.get('/:id', async (req, res) => {
  try {
    const row = await getById(req.params.id);
    if (!row) return res.status(404).json({ error: 'Not found' });
    res.json(row);
  } catch (err) {
    res.status(500).json({ error: err.message });
  }
});

router.post('/', async (req, res) => {
  try {
    const body = req.body;
    const id = body.id || uuidv4();
    const record = { ...body, id };
    const created = await create(record);
    res.status(201).json(created);
  } catch (err) {
    res.status(500).json({ error: err.message });
  }
});

router.put('/:id', async (req, res) => {
  try {
    const id = req.params.id;
    const existing = await getById(id);
    if (!existing) return res.status(404).json({ error: 'Not found' });
    const updated = await update(id, req.body);
    res.json(updated);
  } catch (err) {
    res.status(500).json({ error: err.message });
  }
});

router.delete('/:id', async (req, res) => {
  try {
    const id = req.params.id;
    const existing = await getById(id);
    if (!existing) return res.status(404).json({ error: 'Not found' });
    await remove(id);
    res.status(204).end();
  } catch (err) {
    res.status(500).json({ error: err.message });
  }
});

export default router;
