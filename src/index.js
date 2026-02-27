import express from 'express';
import cors from 'cors';
import { pool } from './db.js';
import dotenv from 'dotenv';

dotenv.config();

const app = express();  // ← PRIMEIRO: declara app
const PORT = process.env.PORT || 3000;

// Middleware
app.use(cors({ origin: true }));  // Permite todos para teste
app.use(express.json());

// Rotas de saúde e debug (DEPOIS do app)
app.get('/', (req, res) => {
  res.json({ ok: true, message: 'API online' });
});

// Health check MySQL
app.get('/health', async (req, res) => {
  try {
    const [rows] = await pool.query('SELECT 1 as test');
    res.json({ 
      status: 'OK', 
      mysql: 'Connected!', 
      test: rows[0].test 
    });
  } catch (err) {
    console.error('MySQL erro:', err.message);
    res.status(500).json({ 
      error: 'MySQL falhou', 
      details: err.message,
      vars: {
        host: !!process.env.MYSQLHOST,
        port: !!process.env.MYSQLPORT,
        user: !!process.env.MYSQLUSER,
        db: !!process.env.MYSQLDATABASE
      }
    });
  }
});

// Debug vars MySQL
app.get('/debug', (req, res) => {
  res.json({
    mysqlVars: {
      host: process.env.MYSQLHOST ? 'OK' : 'MISSING',
      port: process.env.MYSQLPORT,
      user: process.env.MYSQLUSER ? 'OK' : 'MISSING',
      pass: process.env.MYSQLPASSWORD ? 'OK' : 'MISSING',
      db: process.env.MYSQLDATABASE ? 'OK' : 'MISSING'
    }
  });
});

// API Login
app.post('/api/login', async (req, res) => {
  try {
    const { email, senha } = req.body;

    if (!email || !senha) {
      return res.status(400).json({ success: false, message: 'Email e senha são obrigatórios.' });
    }

    const [rows] = await pool.query(
      'SELECT email, nome FROM SF_USUARIO WHERE email = ? AND senha = ? LIMIT 1',
      [email, senha]
    );

    if (rows.length > 0) {
      return res.json({
        success: true,
        message: 'Usuário confirmado com sucesso.',
        email: rows[0].email,
        nome: rows[0].nome,
      });
    } else {
      return res.status(401).json({
        success: false,
        message: 'Email ou senha inválidos.',
      });
    }
  } catch (err) {
    console.error('Erro na rota /api/login:', err);
    res.status(500).json({ success: false, message: 'Erro interno no servidor.', error: err.message });
  }
});


// Inicia servidor
app.listen(PORT, () => {
  console.log(`🚀 API rodando na porta ${PORT}`);
  console.log('✅ Teste: https://sua-url/health');
});

