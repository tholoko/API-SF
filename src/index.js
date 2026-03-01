import express from 'express';
import cors from 'cors';
import { pool } from './db.js';
import dotenv from 'dotenv';
import nodemailer from "nodemailer";
import dns from "node:dns";
dns.setDefaultResultOrder("ipv4first");

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

    if (!sala || !inicio || !fim || !motivo || !usuario) {
      return res.status(400).json({
        success: false,
        message: 'sala, inicio, fim, motivo e usuario são obrigatórios.'
      });
    }

    const ini = new Date(inicio);
    const end = new Date(fim);
    if (!(end > ini)) {
      return res.status(400).json({ success: false, message: 'fim deve ser maior que inicio.' });
    }

    const ids = Array.isArray(participantes)
      ? participantes.map(Number).filter(Number.isFinite)
      : [];

    await conn.beginTransaction();

    // 1) Insere agendamento
    const [ins] = await conn.query(
      `INSERT INTO SF_AGENDAMENTO (sala, inicio, fim, motivo, usuario_agendamento, status, data_agendamento)
       VALUES (?, ?, ?, ?, ?, 'Agendado', NOW())`,
      [sala, inicio, fim, motivo, usuario]
    );

    const idAgendamento = ins.insertId;

    // 2) Carrega usuários convidados
    let convidados = [];
    if (ids.length) {
      const [u] = await conn.query(
        `SELECT id, nome, email
           FROM SF_USUARIO
          WHERE id IN (?)
            AND email IS NOT NULL AND email <> ''`,
        [ids]
      );
      convidados = u;
    }

    // 3) Salva participantes
    for (const p of convidados) {
      await conn.query(
        `INSERT INTO SF_AGENDAMENTO_PARTICIPANTE (id_agendamento, id_usuario, nome, email)
         VALUES (?, ?, ?, ?)`,
        [idAgendamento, p.id, p.nome, p.email]
      );
    }

    // 4) Enfileira convites (sequence=0)
    for (const p of convidados) {
      const uid = `${idAgendamento}-${p.id}@sociedadefranciosi`;

      await conn.query(
        `INSERT INTO SF_EMAIL_QUEUE
          (tipo, status, tentativas, max_tentativas,
           id_agendamento, id_usuario, email, nome,
           sala, inicio, fim, motivo, uid, sequence,
           created_at)
         VALUES
          ('CONVITE_SALA', 'PENDENTE', 0, 5,
           ?, ?, ?, ?,
           ?, ?, ?, ?, ?, 0,
           NOW())`,
        [
          idAgendamento, p.id, p.email, p.nome,
          sala, inicio, fim, motivo, uid
        ]
      );
    }

    await conn.commit();

    if (!convidados.length) {
      return res.json({
        success: true,
        message: 'Agendamento salvo (sem participantes selecionados).',
        id: idAgendamento,
        filaEmail: { total: 0, enfileirados: 0 }
      });
    }

    return res.json({
      success: true,
      message: 'Agendamento salvo. Convites enfileirados para envio.',
      id: idAgendamento,
      filaEmail: { total: convidados.length, enfileirados: convidados.length }
    });
  } catch (err) {
    try { await conn.rollback(); } catch {}
    return res.status(500).json({
      success: false,
      message: 'Erro ao salvar agendamento.',
      error: err.message
    });
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
  const conn = await pool.getConnection();

  try {
    const { id } = req.params;

    const usuarioSolicitante =
      req.headers['x-usuario'] || req.headers['x-user'] || '';

    if (!usuarioSolicitante) {
      return res.status(400).json({ success: false, message: 'Usuário solicitante é obrigatório.' });
    }

    await conn.beginTransaction();

    // 1) Buscar agendamento
    const [agRows] = await conn.query(
      `SELECT id, sala, inicio, fim, motivo, usuario_agendamento, status
         FROM SF_AGENDAMENTO
        WHERE id = ?
        LIMIT 1`,
      [id]
    );

    if (!agRows.length) {
      await conn.rollback();
      return res.status(404).json({ success: false, message: 'Agendamento não encontrado.' });
    }

    const ag = agRows[0];

    if (ag.status !== 'Agendado') {
      await conn.rollback();
      return res.status(409).json({ success: false, message: 'Este agendamento não está mais como Agendado.' });
    }

    if (ag.usuario_agendamento !== usuarioSolicitante) {
      await conn.rollback();
      return res.status(403).json({
        success: false,
        message: 'Você não tem permissão para excluir um agendamento de outro usuário.'
      });
    }

    // 2) Cancelar agendamento
    const [upd] = await conn.query(
      `UPDATE SF_AGENDAMENTO
          SET status = 'Cancelado',
              usuario_cancelamento = ?,
              data_cancelamento = NOW()
        WHERE id = ?
          AND status = 'Agendado'
          AND usuario_agendamento = ?`,
      [usuarioSolicitante, id, usuarioSolicitante]
    );

    if (upd.affectedRows === 0) {
      await conn.rollback();
      return res.status(409).json({ success: false, message: 'Não foi possível cancelar (agendamento já alterado).' });
    }

    // 3) Buscar participantes
    const [parts] = await conn.query(
      `SELECT id_usuario, nome, email
         FROM SF_AGENDAMENTO_PARTICIPANTE
        WHERE id_agendamento = ?
          AND email IS NOT NULL AND email <> ''`,
      [id]
    );

    // 4) Enfileirar cancelamentos (sequence=1) — iTIP pede incremento em CANCEL [web:161]
    for (const p of parts) {
      const uid = `${ag.id}-${p.id_usuario}@sociedadefranciosi`;

      await conn.query(
        `INSERT INTO SF_EMAIL_QUEUE
          (tipo, status, tentativas, max_tentativas,
           id_agendamento, id_usuario, email, nome,
           sala, inicio, fim, motivo, uid, sequence,
           created_at)
         VALUES
          ('CANCELAR_SALA', 'PENDENTE', 0, 5,
           ?, ?, ?, ?,
           ?, ?, ?, ?, ?, 1,
           NOW())`,
        [
          ag.id, p.id_usuario, p.email, p.nome,
          ag.sala, ag.inicio, ag.fim, ag.motivo, uid
        ]
      );
    }

    await conn.commit();

    return res.json({
      success: true,
      message: 'Agendamento cancelado com sucesso. Cancelamentos enfileirados para envio.',
      cancelEmails: { total: parts.length, enfileirados: parts.length }
    });
  } catch (err) {
    try { await conn.rollback(); } catch {}
    console.error('Erro DELETE /api/cancelar-agendamentos/sala/:id:', err);
    return res.status(500).json({ success: false, message: 'Erro interno no servidor.', error: err.message });
  } finally {
    conn.release();
  }
});


app.get('/api/usuarios', async (req, res) => {
  try {
    const [rows] = await pool.query(
      `SELECT id, nome, email, setor
         FROM SF_USUARIO
        WHERE email IS NOT NULL AND email <> ''
        ORDER BY nome ASC`
    );
    res.json({ success: true, items: rows });
  } catch (err) {
    res.status(500).json({ success: false, message: 'Erro ao listar usuários.', error: err.message });
  }
});



// Gmail SMTP (587 STARTTLS)
const transporter = nodemailer.createTransport({
  host: "smtp.gmail.com",
  port: 587,
  secure: false,
  auth: { user: process.env.GMAIL_USER, pass: process.env.GMAIL_APP_PASS },
  tls: { servername: "smtp.gmail.com" },
  // força lookup IPv4
  lookup: (hostname, options, callback) => dns.lookup(hostname, { family: 4 }, callback),
});

// ---------- ICS helpers (os seus, só ajustei CRLF) ----------
function toICSDateUTC(date) {
  const pad = (n) => String(n).padStart(2, "0");
  return (
    date.getUTCFullYear() +
    pad(date.getUTCMonth() + 1) +
    pad(date.getUTCDate()) + "T" +
    pad(date.getUTCHours()) +
    pad(date.getUTCMinutes()) +
    pad(date.getUTCSeconds()) + "Z"
  );
}

function buildICS({ uid, dtstartUtc, dtendUtc, summary, description, location, organizerEmail, attendeeEmail, attendeeName }) {
  // iCal costuma ser mais compatível com CRLF (\r\n)
  const nl = "\r\n";
  return [
    "BEGIN:VCALENDAR",
    "VERSION:2.0",
    "CALSCALE:GREGORIAN",
    "METHOD:REQUEST",
    "PRODID:-//Sociedade Franciosi//Agendamentos//PT-BR",
    "BEGIN:VEVENT",
    `UID:${uid}`,
    `DTSTAMP:${toICSDateUTC(new Date())}`,
    `DTSTART:${toICSDateUTC(dtstartUtc)}`,
    `DTEND:${toICSDateUTC(dtendUtc)}`,
    `SUMMARY:${summary}`,
    `DESCRIPTION:${description}`,
    `LOCATION:${location}`,
    `ORGANIZER:mailto:${organizerEmail}`,
    `ATTENDEE;CN=${attendeeName};RSVP=TRUE:mailto:${attendeeEmail}`,
    "END:VEVENT",
    "END:VCALENDAR",
    "" // final newline
  ].join(nl);
}

// debug env
app.get("/debug-mail", (req, res) => {
  res.json({
    GMAIL_USER: process.env.GMAIL_USER ? "OK" : "MISSING",
    GMAIL_APP_PASS: process.env.GMAIL_APP_PASS ? "OK" : "MISSING",
    MAIL_FROM_NAME: process.env.MAIL_FROM_NAME ? "OK" : "MISSING",
  });
});

// Teste: envia convite
app.get("/api/test-email-startup", async (req, res) => {
  try {
    const toEmail = "lazaro290493@outlook.com";
    const toName = "Lázaro";

    if (!process.env.GMAIL_USER) throw new Error("GMAIL_USER ausente");
    if (!process.env.GMAIL_APP_PASS) throw new Error("GMAIL_APP_PASS ausente");

    await transporter.verify(); // valida credenciais/SMTP

    const uid = `teste-${Date.now()}@gmail.com`;
    const inicio = new Date(Date.now() + 60 * 60 * 1000);
    const fim = new Date(Date.now() + 2 * 60 * 60 * 1000);

    const icsText = buildICS({
      uid,
      dtstartUtc: inicio,
      dtendUtc: fim,
      summary: "Teste convite (Gmail SMTP)",
      description: "Convite de teste enviado via Nodemailer + Gmail.",
      location: "Sala 01",
      organizerEmail: process.env.GMAIL_USER,
      attendeeEmail: toEmail,
      attendeeName: toName,
    });

    const info = await transporter.sendMail({
      from: `"${process.env.MAIL_FROM_NAME || "Sociedade Franciosi"}" <${process.env.GMAIL_USER}>`,
      to: `"${toName}" <${toEmail}>`,
      subject: "Teste convite (ICS via Gmail)",
      text: "Segue convite de calendário.",
      html: "<p>Segue convite de calendário.</p>",
      icalEvent: {
        filename: "invitation.ics",
        method: "REQUEST",
        content: icsText,
      },
    });

    return res.json({ success: true, to: toEmail, messageId: info.messageId, accepted: info.accepted });
  } catch (e) {
    return res.status(500).json({ success: false, error: e?.message || String(e) });
  }
});


app.get('/api/setores', async (req, res) => {
  try {
    const [rows] = await pool.query(
      `SELECT DISTINCT nome 
         FROM SF_SETOR 
        WHERE nome IS NOT NULL AND nome <> ''
        ORDER BY nome ASC`
    );
    res.json({ success: true, items: rows });
  } catch (err) {
    res.status(500).json({ success: false, message: 'Erro ao listar setores.', error: err.message });
  }
});

app.post('/api/gestao/usuarios', async (req, res) => {
  const nome = titleCaseNome(req.body?.nome);
  const email = normalizarEmail(req.body?.email);
  const senha = (req.body?.senha || '').toString();
  const telefone = somenteNumeros(req.body?.telefone);
  const perfil = (req.body?.perfil || '').toString().trim(); // nome do perfil
  const setor = (req.body?.setor || '').toString().trim();   // nome do setor
  const status = (req.body?.status || 'Ativo').toString().trim();

  if (!nome) return res.status(400).json({ error: 'Nome é obrigatório.' });
  if (!email) return res.status(400).json({ error: 'Email é obrigatório.' });
  if (!senha || senha.length < 6) return res.status(400).json({ error: 'Senha inválida (mínimo 6).' });
  if (!perfil) return res.status(400).json({ error: 'Perfil é obrigatório.' });
  if (!setor) return res.status(400).json({ error: 'Setor é obrigatório.' });

  const saltRounds = 12;
  const senhaHash = await bcrypt.hash(senha, saltRounds); // armazena hash na coluna SENHA [web:83]

  const [r] = await db.query(
    `INSERT INTO SF_USUARIO (NOME, EMAIL, SENHA, TELEFONE, PERFIL, SETOR, STATUS)
     VALUES (?, ?, ?, ?, ?, ?, ?)`,
    [nome, email, senhaHash, telefone, perfil, setor, status]
  );

  res.status(201).json({ id: r.insertId, nome, email, perfil, setor, status });
});

app.get('/api/gestao/usuarios', async (req, res) => {
  const [rows] = await db.query(
    `SELECT ID, NOME, EMAIL, SETOR, STATUS
     FROM SF_USUARIO
     ORDER BY NOME`
  );

  res.json(rows.map(r => ({
    id: r.ID,
    nome: r.NOME,
    email: r.EMAIL,
    setor: r.SETOR,
    status: r.STATUS,
  })));
});

app.get('/api/gestao/usuarios/:id', async (req, res) => {
  const id = Number(req.params.id);
  if (!id) return res.status(400).json({ error: 'ID inválido.' });

  const [rows] = await db.query(
    `SELECT ID, NOME, EMAIL, TELEFONE, PERFIL, SETOR, STATUS
     FROM SF_USUARIO
     WHERE ID = ?`,
    [id]
  );

  if (!rows.length) return res.status(404).json({ error: 'Usuário não encontrado.' });
  res.json(rows[0]);
});

app.put('/api/gestao/usuarios/:id', async (req, res) => {
  const id = Number(req.params.id);
  if (!id) return res.status(400).json({ error: 'ID inválido.' });

  const nome = titleCaseNome(req.body?.nome);
  const email = normalizarEmail(req.body?.email);
  const telefone = somenteNumeros(req.body?.telefone);
  const perfil = (req.body?.perfil || '').toString().trim();
  const setor = (req.body?.setor || '').toString().trim();
  const status = (req.body?.status || '').toString().trim();

  if (!nome) return res.status(400).json({ error: 'Nome é obrigatório.' });
  if (!email) return res.status(400).json({ error: 'Email é obrigatório.' });
  if (!perfil) return res.status(400).json({ error: 'Perfil é obrigatório.' });
  if (!setor) return res.status(400).json({ error: 'Setor é obrigatório.' });
  if (!status) return res.status(400).json({ error: 'Status é obrigatório.' });

  const [r] = await db.query(
    `UPDATE SF_USUARIO
     SET NOME = ?, EMAIL = ?, TELEFONE = ?, PERFIL = ?, SETOR = ?, STATUS = ?
     WHERE ID = ?`,
    [nome, email, telefone, perfil, setor, status, id]
  );

  if (r.affectedRows === 0) return res.status(404).json({ error: 'Usuário não encontrado.' });
  res.json({ ok: true });
});

app.patch('/api/gestao/usuarios/:id/status', async (req, res) => {
  const id = Number(req.params.id);
  const status = (req.body?.status || '').toString().trim();
  if (!id) return res.status(400).json({ error: 'ID inválido.' });
  if (!status) return res.status(400).json({ error: 'Status é obrigatório.' });

  const [r] = await db.query(`UPDATE SF_USUARIO SET STATUS = ? WHERE ID = ?`, [status, id]);
  if (r.affectedRows === 0) return res.status(404).json({ error: 'Usuário não encontrado.' });

  res.json({ ok: true });
});

app.delete('/api/gestao/usuarios/:id', async (req, res) => {
  const id = Number(req.params.id);
  if (!id) return res.status(400).json({ error: 'ID inválido.' });

  const [r] = await db.query(`DELETE FROM SF_USUARIO WHERE ID = ?`, [id]);
  if (r.affectedRows === 0) return res.status(404).json({ error: 'Usuário não encontrado.' });

  res.json({ ok: true });
});

app.get('/api/gestao/usuarios/perfis', async (req, res) => {
  const [rows] = await db.query(`SELECT ID, NOME FROM SF_PERFIL ORDER BY NOME`);
  res.json(rows);
});

app.get('/api/gestao/usuarios/setores', async (req, res) => {
  const [rows] = await db.query(`SELECT ID, NOME FROM SF_SETOR ORDER BY NOME`);
  res.json(rows);
});

app.post('/api/gestao/usuarios/setores', async (req, res) => {
  const nome = titleCaseNome(req.body?.nome);
  if (!nome) return res.status(400).json({ error: 'Nome do setor é obrigatório.' });

  const [r] = await db.query(`INSERT INTO SF_SETOR (NOME) VALUES (?)`, [nome]);
  res.status(201).json({ id: r.insertId, nome });
});

app.patch('/api/gestao/usuarios/:id/senha', async (req, res) => {
  const id = Number(req.params.id);
  const senhaAtual = (req.body?.senhaAtual || '').toString();
  const novaSenha = (req.body?.novaSenha || '').toString();

  if (!id) return res.status(400).json({ error: 'ID inválido.' });
  if (!senhaAtual) return res.status(400).json({ error: 'senhaAtual é obrigatória.' });
  if (!novaSenha || novaSenha.length < 6) {
    return res.status(400).json({ error: 'novaSenha inválida (mínimo 6 caracteres).' });
  }

  const [rows] = await db.query(`SELECT SENHA FROM SF_USUARIO WHERE ID = ?`, [id]);
  if (!rows.length) return res.status(404).json({ error: 'Usuário não encontrado.' });

  const hashAtual = rows[0].SENHA;

  const ok = await bcrypt.compare(senhaAtual, hashAtual); // compara texto com hash [web:83]
  if (!ok) return res.status(401).json({ error: 'Senha atual incorreta.' });

  const saltRounds = 12;
  const novoHash = await bcrypt.hash(novaSenha, saltRounds); // gera novo hash [web:83]

  await db.query(`UPDATE SF_USUARIO SET SENHA = ? WHERE ID = ?`, [novoHash, id]);

  res.json({ ok: true });
});

app.patch('/api/gestao/usuarios/:id/senha-reset', async (req, res) => {
  const id = Number(req.params.id);
  const novaSenha = (req.body?.novaSenha || '').toString();

  if (!id) return res.status(400).json({ error: 'ID inválido.' });
  if (!novaSenha || novaSenha.length < 6) {
    return res.status(400).json({ error: 'novaSenha inválida (mínimo 6 caracteres).' });
  }

  const saltRounds = 12;
  const novoHash = await bcrypt.hash(novaSenha, saltRounds); // hash bcrypt [web:83]

  const [r] = await db.query(`UPDATE SF_USUARIO SET SENHA = ? WHERE ID = ?`, [novoHash, id]);
  if (r.affectedRows === 0) return res.status(404).json({ error: 'Usuário não encontrado.' });

  res.json({ ok: true });
});

// Inicia servidor
app.listen(PORT, () => {
  console.log(`🚀 API rodando na porta ${PORT}`);
  console.log('✅ Teste: https://sua-url/health');
});

