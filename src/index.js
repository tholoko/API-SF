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

(async () => {
  try {
    await pool.query("SET time_zone = '-03:00'"); // ajusta a sessão [web:209]
    console.log('MySQL time_zone ajustado para -03:00');
  } catch (e) {
    console.error('Falha ao setar time_zone:', e);
  }
})();

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

app.post('/api/agendamentos/sala/verificar', async (req, res) => {
  try {
    const { sala, inicio, fim } = req.body;

    if (!sala || !inicio || !fim) {
      return res.status(400).json({ success: false, message: 'sala, inicio e fim são obrigatórios.' });
    }

    const ini = new Date(inicio);
    const end = new Date(fim);
    if (!(end > ini)) {
      return res.status(400).json({ success: false, message: 'fim deve ser maior que inicio.' });
    }

    const [rows] = await pool.query(
      `
      SELECT
        sala,
        inicio,
        fim,
        motivo,
        usuario_agendamento,
        data_agendamento
      FROM SF_AGENDAMENTO
      WHERE sala = ?
        AND status = 'Agendado'
        AND inicio < ?
        AND fim > ?
      ORDER BY inicio ASC
      LIMIT 1
      `,
      [sala, fim, inicio]
    );

    if (rows.length > 0) {
      return res.json({
        success: true,
        conflito: true,
        message: 'Existe conflito de agendamento.',
        conflitoDetalhe: rows[0]
      });
    }

    return res.json({ success: true, conflito: false, message: 'Sem conflito.' });
  } catch (err) {
    console.error('Erro /api/agendamentos/sala/verificar:', err);
    return res.status(500).json({ success: false, message: 'Erro interno no servidor.', error: err.message });
  }
});

app.post('/api/agendamentos/sala', async (req, res) => {
  const conn = await pool.getConnection();
  try {
    const { sala, inicio, fim, motivo, usuario, participantes } = req.body;

    // participantes: array de { id } OU array de ids. (você escolhe no front)
    const ids = Array.isArray(participantes) ? participantes : [];

    await conn.beginTransaction();

    // 1) insert agendamento (ajuste colunas conforme seu schema)
    const [ins] = await conn.query(
      `INSERT INTO SF_AGENDAMENTO (sala, inicio, fim, motivo, usuario_agendamento, status, data_agendamento)
       VALUES (?, ?, ?, ?, ?, 'Agendado', NOW())`,
      [sala, inicio, fim, motivo, usuario]
    );

    const idAgendamento = ins.insertId;

    // 2) carrega usuários convidados
    let convidados = [];
    if (ids.length) {
      const [u] = await conn.query(
        `SELECT id, nome, email FROM SF_USUARIO WHERE id IN (?)`,
        [ids]
      );
      convidados = u;
    }

    // 3) salva participantes
    for (const p of convidados) {
      await conn.query(
        `INSERT INTO SF_AGENDAMENTO_PARTICIPANTE (id_agendamento, id_usuario, nome, email)
         VALUES (?, ?, ?, ?)`,
        [idAgendamento, p.id, p.nome, p.email]
      );
    }

    await conn.commit();

    // 4) enviar convites (fora da transaction)
    // ATENÇÃO: inicio/fim chegam como datetime-local ("YYYY-MM-DDTHH:mm"), trate como hora local do Brasil.
    // Aqui fica simples: converter para Date e depois UTC (depende da sua regra de timezone).
    const dtStart = new Date(inicio);
    const dtEnd = new Date(fim);

    for (const p of convidados) {
      const uid = `${idAgendamento}-${p.id}@sociedadefranciosi`;
      const ics = buildICS({
        uid,
        dtstartUtc: new Date(dtStart.toISOString()),
        dtendUtc: new Date(dtEnd.toISOString()),
        summary: `Reunião - ${sala}`,
        description: motivo || '',
        location: sala,
        organizerEmail: process.env.GMAIL_USER,
        attendeeEmail: p.email,
        attendeeName: p.nome
      });

      await transporter.sendMail({
        from: `"Sociedade Franciosi" <${process.env.GMAIL_USER}>`,
        to: p.email,
        subject: `Convite: Reunião (${sala})`,
        text: `Você foi convidado para uma reunião.\nSala: ${sala}\nInício: ${inicio}\nFim: ${fim}\nMotivo: ${motivo}`,
        icalEvent: {
          filename: 'convite.ics',
          method: 'REQUEST',
          content: ics
        }
      });
    }

    return res.json({ success: true, message: 'Agendamento salvo e convites enviados.', id: idAgendamento });
  } catch (err) {
    try { await conn.rollback(); } catch {}
    return res.status(500).json({ success: false, message: 'Erro ao salvar agendamento.', error: err.message });
  } finally {
    conn.release();
  }
});


app.get('/api/agendamentos/sala/dia', async (req, res) => {
  try {
    const { data } = req.query; // opcional: '2026-02-28'
    // se não vier, pega hoje no banco (server time)
    const [rows] = await pool.query(
      `
      SELECT
        id,
        sala,
        inicio,
        fim,
        motivo,
        usuario_agendamento,
        data_agendamento
      FROM SF_AGENDAMENTO
      WHERE status = 'Agendado'
        AND DATE(inicio) = COALESCE(?, CURDATE())
      ORDER BY inicio ASC
      `,
      [data || null]
    );

    return res.json({ success: true, items: rows });
  } catch (err) {
    console.error('Erro /api/agendamentos/sala/dia:', err);
    res.status(500).json({ success: false, message: 'Erro interno no servidor.', error: err.message });
  }
});

app.delete('/api/cancelar-agendamentos/sala/:id', async (req, res) => {
  try {
    const { id } = req.params;

    const usuarioSolicitante =
      req.headers['x-usuario'] || req.headers['x-user'] || '';

    if (!usuarioSolicitante) {
      return res.status(400).json({ success: false, message: 'Usuário solicitante é obrigatório.' });
    }

    // Cancela SOMENTE se:
    // - existir
    // - status = Agendado
    // - o usuário solicitante for o mesmo do usuario_agendamento
    // E grava usuário/data do cancelamento
    const [result] = await pool.query(
      `UPDATE SF_AGENDAMENTO
          SET status = 'Cancelado',
              usuario_cancelamento = ?,
              data_cancelamento = NOW()
        WHERE id = ?
          AND status = 'Agendado'
          AND usuario_agendamento = ?`,
      [usuarioSolicitante, id, usuarioSolicitante]
    );

    if (result.affectedRows > 0) {
      return res.json({ success: true, message: 'Agendamento cancelado com sucesso.' });
    }

    // Se não alterou, descobrir se existe e se é de outro usuário
    const [rows] = await pool.query(
      `SELECT id, usuario_agendamento, status
         FROM SF_AGENDAMENTO
        WHERE id = ?
        LIMIT 1`,
      [id]
    );

    if (rows.length === 0) {
      return res.status(404).json({ success: false, message: 'Agendamento não encontrado.' });
    }

    if (rows[0].status !== 'Agendado') {
      return res.status(409).json({ success: false, message: 'Este agendamento não está mais como Agendado.' });
    }

    return res.status(403).json({
      success: false,
      message: 'Você não tem permissão para excluir um agendamento de outro usuário.'
    });
  } catch (err) {
    console.error('Erro DELETE /api/cancelar-agendamentos/sala/:id:', err);
    return res.status(500).json({ success: false, message: 'Erro interno no servidor.', error: err.message });
  }
});

app.get('/api/usuarios', async (req, res) => {
  try {
    const [rows] = await pool.query(
      `SELECT id, nome, email
         FROM SF_USUARIO
        WHERE email IS NOT NULL AND email <> ''
        ORDER BY nome ASC`
    );
    res.json({ success: true, items: rows });
  } catch (err) {
    res.status(500).json({ success: false, message: 'Erro ao listar usuários.', error: err.message });
  }
});

import nodemailer from 'nodemailer';

const transporter = nodemailer.createTransport({
  service: 'gmail',
  auth: {
    user: process.env.GMAIL_USER,      // seuemail@gmail.com
    pass: process.env.GMAIL_APP_PASS   // app password (não é a senha normal) [web:445]
  }
});

function toICSDateUTC(date) {
  // date é um Date em UTC; formato YYYYMMDDTHHMMSSZ
  const pad = (n) => String(n).padStart(2, '0');
  return (
    date.getUTCFullYear() +
    pad(date.getUTCMonth() + 1) +
    pad(date.getUTCDate()) + 'T' +
    pad(date.getUTCHours()) +
    pad(date.getUTCMinutes()) +
    pad(date.getUTCSeconds()) + 'Z'
  );
}

function buildICS({ uid, dtstartUtc, dtendUtc, summary, description, location, organizerEmail, attendeeEmail, attendeeName }) {
  // básico e compatível com a maioria dos clientes
  return `BEGIN:VCALENDAR
  VERSION:2.0
  CALSCALE:GREGORIAN
  METHOD:REQUEST
  PRODID:-//Sociedade Franciosi//Agendamentos//PT-BR
  BEGIN:VEVENT
  UID:${uid}
  DTSTAMP:${toICSDateUTC(new Date())}
  DTSTART:${toICSDateUTC(dtstartUtc)}
  DTEND:${toICSDateUTC(dtendUtc)}
  SUMMARY:${summary}
  DESCRIPTION:${description}
  LOCATION:${location}
  ORGANIZER:mailto:${organizerEmail}
  ATTENDEE;CN=${attendeeName};RSVP=TRUE:mailto:${attendeeEmail}
  END:VEVENT
  END:VCALENDAR`;
}

// Inicia servidor
app.listen(PORT, () => {
  console.log(`🚀 API rodando na porta ${PORT}`);
  console.log('✅ Teste: https://sua-url/health');
});

