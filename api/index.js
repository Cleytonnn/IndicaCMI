import express from 'express';
import cors from 'cors';
import recordsRouter from '../server/routes/records.js';

const app = express();

app.use(cors());
app.use(express.json({ limit: '10mb' }));

app.use('/api/records', recordsRouter);

app.get('/api', (req, res) => {
  res.json({ status: 'ok' });
});

export default app;
