import express from 'express';
import cors from 'cors';
import { pool } from './db.js';
import dotenv from 'dotenv';

dotenv.config();

const app = express();
const PORT = process.env.PORT || 3000;

app.use(cors());
app.use(express.json());

// Rota de saúde
app.get('/', (req, res) => {
  res.json({ ok: true, message: 'API online' });
});

// Rota de login: valida SF_USUARIO (email + senha)
app.post('/api/login', async (req, res) => {
  try {
    const { email, senha } = req.body;

    if (!email || !senha) {
      return res.status(400).json({ success: false, message: 'Email e senha são obrigatórios.' });
    }

    const [rows] = await pool.query(
      'SELECT email FROM SF_USUARIO WHERE email = ? AND senha = ? LIMIT 1',
      [email, senha]
    );

    if (rows.length > 0) {
      return res.json({
        success: true,
        message: 'Usuário confirmado com sucesso.',
        email: rows[0].email,
      });
    } else {
      return res.status(401).json({
        success: false,
        message: 'Email ou senha inválidos.',
      });
    }
  } catch (err) {
    console.error('Erro na rota /api/login:', err);
    res.status(500).json({ success: false, message: 'Erro interno no servidor.' });
  }
});

app.listen(PORT, () => {
  console.log(`API rodando na porta ${PORT}`);
});
