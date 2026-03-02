import express from 'express';
import cors from 'cors';
import { pool } from './db.js';
import dotenv from 'dotenv';
import dns from "node:dns";
import bcrypt from 'bcryptjs';
import { titleCaseNome, normalizarEmail, somenteNumeros } from './utils.js';
import crypto from 'node:crypto';

import fs from "node:fs";
import path from "node:path";
import multer from "multer";

dns.setDefaultResultOrder("ipv4first");
dotenv.config();

const app = express();
const PORT = process.env.PORT || 3000;

// =====================
// Middleware base
// =====================
app.use(cors({ origin: true }));
app.use(express.json());

// =====================
// Ajuste timezone MySQL
// =====================
(async () => {
  try {
    await pool.query("SET time_zone = '-03:00'");
    console.log('MySQL time_zone ajustado para -03:00');
  } catch (e) {
    console.error('Falha ao setar time_zone:', e);
  }
})();

// =====================
// Rotas de saúde / debug
// =====================
app.get('/', (req, res) => {
  res.json({ ok: true, message: 'API online' });
});

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

// =====================
// API Login
// =====================
app.post('/api/login', async (req, res) => {
  try {
    const email = normalizarEmail(req.body?.email);
    const senha = (req.body?.senha || '').toString();

    if (!email || !senha) {
      return res.status(400).json({ success: false, message: 'Email e senha são obrigatórios.' });
    }

    const [rows] = await pool.query(
      `SELECT ID, EMAIL, NOME, SENHA, STATUS, MUST_CHANGE_PASSWORD
         FROM SF_USUARIO
        WHERE EMAIL = ?
        LIMIT 1`,
      [email]
    );

    if (!rows.length) {
      return res.status(401).json({ success: false, message: 'Email ou senha inválidos.' });
    }

    const u = rows[0];

    if ((u.STATUS || '').toString().trim() !== 'Ativo') {
      return res.status(403).json({ success: false, message: 'Usuário desativado.' });
    }

    const ok = await bcrypt.compare(senha, u.SENHA);
    if (!ok) {
      return res.status(401).json({ success: false, message: 'Email ou senha inválidos.' });
    }

    return res.json({
      success: true,
      email: u.EMAIL,
      nome: u.NOME,
      id: u.ID,
      mustChangePassword: Number(u.MUST_CHANGE_PASSWORD) === 1
    });
  } catch (err) {
    console.error('Erro /api/login:', err);
    return res.status(500).json({ success: false, message: 'Erro interno.', error: err.message });
  }
});

app.post('/api/usuarios/primeiro-acesso/senha', async (req, res) => {
  try {
    const email = normalizarEmail(req.body?.email);
    const newPassword = (req.body?.newPassword || '').toString();

    if (!email || !newPassword) return res.status(400).json({ success: false, message: 'Dados incompletos.' });
    if (newPassword.length < 6) return res.status(400).json({ success: false, message: 'Senha mínima: 6 caracteres.' });

    const hash = await bcrypt.hash(newPassword, 12);

    await pool.query(
      `UPDATE SF_USUARIO
          SET SENHA = ?, MUST_CHANGE_PASSWORD = 0
        WHERE EMAIL = ?`,
      [hash, email]
    );

    return res.json({ success: true });
  } catch (err) {
    console.error('Erro primeiro acesso senha:', err);
    return res.status(500).json({ success: false, message: 'Erro ao atualizar senha.', error: err.message });
  }
});

// =====================
// Agendamentos - Sala
// =====================
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

    const [ins] = await conn.query(
      `INSERT INTO SF_AGENDAMENTO (sala, inicio, fim, motivo, usuario_agendamento, status, data_agendamento)
       VALUES (?, ?, ?, ?, ?, 'Agendado', NOW())`,
      [sala, inicio, fim, motivo, usuario]
    );

    const idAgendamento = ins.insertId;

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

    for (const p of convidados) {
      await conn.query(
        `INSERT INTO SF_AGENDAMENTO_PARTICIPANTE (id_agendamento, id_usuario, nome, email)
         VALUES (?, ?, ?, ?)`,
        [idAgendamento, p.id, p.nome, p.email]
      );
    }

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
    const { data } = req.query;
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

    const [parts] = await conn.query(
      `SELECT id_usuario, nome, email
         FROM SF_AGENDAMENTO_PARTICIPANTE
        WHERE id_agendamento = ?
          AND email IS NOT NULL AND email <> ''`,
      [id]
    );

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

// =====================
// Usuários / Setores
// =====================
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

// =====================
// Gestão Usuários
// =====================
app.get('/api/gestao-usuarios-perfis', async (req, res) => {
  try {
    const [rows] = await pool.query(
      `SELECT ID, NOME
         FROM SF_PERFIL
        WHERE NOME IS NOT NULL AND NOME <> ''
        ORDER BY NOME ASC`
    );
    res.json({ success: true, items: rows });
  } catch (err) {
    res.status(500).json({ success: false, message: 'Erro ao listar perfis.', error: err.message });
  }
});

app.get('/api/gestao-usuarios-setores', async (req, res) => {
  try {
    const [rows] = await pool.query(
      `SELECT ID, NOME
         FROM SF_SETOR
        WHERE NOME IS NOT NULL AND NOME <> ''
        ORDER BY NOME ASC`
    );
    res.json({ success: true, items: rows });
  } catch (err) {
    res.status(500).json({ success: false, message: 'Erro ao listar setores.', error: err.message });
  }
});

app.post('/api/gestao-usuarios-setores', async (req, res) => {
  try {
    const nome = titleCaseNome(req.body?.nome);
    if (!nome) return res.status(400).json({ success: false, message: 'Nome do setor é obrigatório.' });

    const [r] = await pool.query(`INSERT INTO SF_SETOR (NOME) VALUES (?)`, [nome]);
    res.status(201).json({ success: true, item: { id: r.insertId, nome } });
  } catch (err) {
    res.status(500).json({ success: false, message: 'Erro ao adicionar setor.', error: err.message });
  }
});

app.get('/api/gestao-usuarios', async (req, res) => {
  try {
    const [rows] = await pool.query(
      `SELECT ID, NOME, EMAIL, SETOR, STATUS
         FROM SF_USUARIO
        ORDER BY NOME ASC`
    );

    const items = rows.map(r => ({
      id: r.ID,
      nome: r.NOME,
      email: r.EMAIL,
      setor: r.SETOR,
      status: r.STATUS,
    }));

    res.json({ success: true, items });
  } catch (err) {
    res.status(500).json({ success: false, message: 'Erro ao listar usuários.', error: err.message });
  }
});

app.post('/api/gestao-usuarios-adicionar', async (req, res) => {
  try {
    const nome = titleCaseNome(req.body?.nome);
    const email = normalizarEmail(req.body?.email);
    const senha = (req.body?.senha || '').toString();
    const telefone = somenteNumeros(req.body?.telefone);
    const perfil = (req.body?.perfil || '').toString().trim();
    const setor = (req.body?.setor || '').toString().trim();
    const status = (req.body?.status || 'Ativo').toString().trim();

    if (!nome) return res.status(400).json({ success: false, message: 'Nome é obrigatório.' });
    if (!email) return res.status(400).json({ success: false, message: 'Email é obrigatório.' });
    if (!senha || senha.length < 6) return res.status(400).json({ success: false, message: 'Senha inválida (mínimo 6).' });
    if (!perfil) return res.status(400).json({ success: false, message: 'Perfil é obrigatório.' });
    if (!setor) return res.status(400).json({ success: false, message: 'Setor é obrigatório.' });

    const saltRounds = 12;
    const senhaHash = await bcrypt.hash(senha, saltRounds);

    const [r] = await pool.query(
      `INSERT INTO SF_USUARIO (NOME, EMAIL, SENHA, TELEFONE, PERFIL, SETOR, STATUS, MUST_CHANGE_PASSWORD)
       VALUES (?, ?, ?, ?, ?, ?, ?, 1)`,
      [nome, email, senhaHash, telefone, perfil, setor, status]
    );

    res.status(201).json({
      success: true,
      item: { id: r.insertId, nome, email, telefone, perfil, setor, status },
    });
  } catch (err) {
    console.error('ERRO /api/gestao-usuarios-adicionar:', err);
    res.status(500).json({ success: false, message: 'Erro ao criar usuário.', error: err.message });
  }
});

app.get('/api/gestao-usuarios/:id(\\d+)', async (req, res) => {
  try {
    const id = Number(req.params.id);

    const [rows] = await pool.query(
      `SELECT ID, NOME, EMAIL, TELEFONE, PERFIL, SETOR, STATUS
         FROM SF_USUARIO
        WHERE ID = ?`,
      [id]
    );

    if (!rows.length) return res.status(404).json({ success: false, message: 'Usuário não encontrado.' });
    res.json({ success: true, item: rows[0] });
  } catch (err) {
    res.status(500).json({ success: false, message: 'Erro ao buscar usuário.', error: err.message });
  }
});

app.put('/api/gestao-usuarios/:id(\\d+)', async (req, res) => {
  try {
    const id = Number(req.params.id);

    const nome = titleCaseNome(req.body?.nome);
    const email = normalizarEmail(req.body?.email);
    const telefone = somenteNumeros(req.body?.telefone);
    const perfil = (req.body?.perfil || '').toString().trim();
    const setor = (req.body?.setor || '').toString().trim();
    const status = (req.body?.status || '').toString().trim();

    if (!nome) return res.status(400).json({ success: false, message: 'Nome é obrigatório.' });
    if (!email) return res.status(400).json({ success: false, message: 'Email é obrigatório.' });
    if (!perfil) return res.status(400).json({ success: false, message: 'Perfil é obrigatório.' });
    if (!setor) return res.status(400).json({ success: false, message: 'Setor é obrigatório.' });
    if (!status) return res.status(400).json({ success: false, message: 'Status é obrigatório.' });

    const [r] = await pool.query(
      `UPDATE SF_USUARIO
          SET NOME = ?, EMAIL = ?, TELEFONE = ?, PERFIL = ?, SETOR = ?, STATUS = ?
        WHERE ID = ?`,
      [nome, email, telefone, perfil, setor, status, id]
    );

    if (r.affectedRows === 0) return res.status(404).json({ success: false, message: 'Usuário não encontrado.' });
    res.json({ success: true });
  } catch (err) {
    res.status(500).json({ success: false, message: 'Erro ao atualizar usuário.', error: err.message });
  }
});

app.patch('/api/gestao-usuarios/:id(\\d+)/status', async (req, res) => {
  try {
    const id = Number(req.params.id);
    const status = (req.body?.status || '').toString().trim();

    if (!status) return res.status(400).json({ success: false, message: 'Status é obrigatório.' });

    const [r] = await pool.query(`UPDATE SF_USUARIO SET STATUS = ? WHERE ID = ?`, [status, id]);
    if (r.affectedRows === 0) return res.status(404).json({ success: false, message: 'Usuário não encontrado.' });

    res.json({ success: true });
  } catch (err) {
    res.status(500).json({ success: false, message: 'Erro ao alterar status.', error: err.message });
  }
});

app.delete('/api/gestao-usuarios/:id(\\d+)', async (req, res) => {
  try {
    const id = Number(req.params.id);

    const [r] = await pool.query(`DELETE FROM SF_USUARIO WHERE ID = ?`, [id]);
    if (r.affectedRows === 0) return res.status(404).json({ success: false, message: 'Usuário não encontrado.' });

    res.json({ success: true });
  } catch (err) {
    res.status(500).json({ success: false, message: 'Erro ao excluir usuário.', error: err.message });
  }
});

app.patch('/api/gestao-usuarios/:id(\\d+)/senha', async (req, res) => {
  try {
    const id = Number(req.params.id);
    const senhaAtual = (req.body?.senhaAtual || '').toString();
    const novaSenha = (req.body?.novaSenha || '').toString();

    if (!senhaAtual) return res.status(400).json({ success: false, message: 'senhaAtual é obrigatória.' });
    if (!novaSenha || novaSenha.length < 6) return res.status(400).json({ success: false, message: 'novaSenha inválida (mínimo 6).' });

    const [rows] = await pool.query(`SELECT SENHA FROM SF_USUARIO WHERE ID = ?`, [id]);
    if (!rows.length) return res.status(404).json({ success: false, message: 'Usuário não encontrado.' });

    const ok = await bcrypt.compare(senhaAtual, rows[0].SENHA);
    if (!ok) return res.status(401).json({ success: false, message: 'Senha atual incorreta.' });

    const novoHash = await bcrypt.hash(novaSenha, 12);
    await pool.query(`UPDATE SF_USUARIO SET SENHA = ? WHERE ID = ?`, [novoHash, id]);

    res.json({ success: true });
  } catch (err) {
    res.status(500).json({ success: false, message: 'Erro ao trocar senha.', error: err.message });
  }
});

app.patch('/api/gestao-usuarios/:id(\\d+)/senha-reset', async (req, res) => {
  try {
    const id = Number(req.params.id);
    const novaSenha = (req.body?.novaSenha || '').toString();

    if (!novaSenha || novaSenha.length < 6) return res.status(400).json({ success: false, message: 'novaSenha inválida (mínimo 6).' });

    const novoHash = await bcrypt.hash(novaSenha, 12);

    const [r] = await pool.query(`UPDATE SF_USUARIO SET SENHA = ? WHERE ID = ?`, [novoHash, id]);
    if (r.affectedRows === 0) return res.status(404).json({ success: false, message: 'Usuário não encontrado.' });

    res.json({ success: true });
  } catch (err) {
    res.status(500).json({ success: false, message: 'Erro ao resetar senha.', error: err.message });
  }
});

// =====================
// Password reset
// =====================
app.post('/api/password-reset/confirm', async (req, res) => {
  try {
    const email = normalizarEmail(req.body?.email);
    const token = (req.body?.token || '').toString().trim();
    const newPassword = (req.body?.newPassword || '').toString();

    if (!email || !token || !newPassword) {
      return res.status(400).json({ success: false, message: 'Dados incompletos.' });
    }
    if (newPassword.length < 6) {
      return res.status(400).json({ success: false, message: 'Senha mínima: 6 caracteres.' });
    }

    const [rows] = await pool.query(
      `SELECT id, token_hash, expires_at
         FROM SF_PASSWORD_RESET
        WHERE email = ? AND token_hash IS NOT NULL
        ORDER BY id DESC
        LIMIT 1`,
      [email]
    );

    if (!rows.length) return res.status(400).json({ success: false, message: 'Token inválido.' });

    const r = rows[0];

    const [exp] = await pool.query(
      `SELECT (UTC_TIMESTAMP() <= ?) AS ok`,
      [r.expires_at]
    );
    if (!exp?.length || exp[0].ok !== 1) {
      return res.status(400).json({ success: false, message: 'Token expirado.' });
    }

    const ok = await bcrypt.compare(token, r.token_hash);
    if (!ok) return res.status(400).json({ success: false, message: 'Token inválido.' });

    const senhaHash = await bcrypt.hash(newPassword, 12);
    await pool.query(`UPDATE SF_USUARIO SET SENHA = ? WHERE EMAIL = ?`, [senhaHash, email]);

    return res.json({ success: true });
  } catch (err) {
    console.error('password-reset/confirm:', err);
    return res.status(500).json({ success: false, message: 'Erro ao atualizar senha.', error: err.message });
  }
});

app.post('/api/password-reset/verify', async (req, res) => {
  try {
    const email = normalizarEmail(req.body?.email);
    const code = (req.body?.code || '').toString().trim();
    if (!email || !code) return res.status(400).json({ success: false, message: 'Email e código são obrigatórios.' });

    const [rows] = await pool.query(
      `SELECT id, code_hash, expires_at, attempts
         FROM SF_PASSWORD_RESET
        WHERE email = ? AND used = 0
        ORDER BY id DESC
        LIMIT 1`,
      [email]
    );

    if (!rows.length) return res.status(400).json({ success: false, message: 'Código inválido ou já utilizado.' });

    const r = rows[0];

    const [exp] = await pool.query(
      `SELECT (UTC_TIMESTAMP() <= ?) AS ok`,
      [r.expires_at]
    );
    if (!exp?.length || exp[0].ok !== 1) {
      return res.status(400).json({ success: false, message: 'Código expirado.' });
    }

    const ok = await bcrypt.compare(code, r.code_hash);
    if (!ok) {
      await pool.query(`UPDATE SF_PASSWORD_RESET SET attempts = attempts + 1 WHERE id = ?`, [r.id]);
      return res.status(400).json({ success: false, message: 'Código inválido.' });
    }

    const token = crypto.randomBytes(24).toString('hex');
    const tokenHash = await bcrypt.hash(token, 10);

    await pool.query(
      `UPDATE SF_PASSWORD_RESET
          SET token_hash = ?, used = 1
        WHERE id = ?`,
      [tokenHash, r.id]
    );

    return res.json({ success: true, token });
  } catch (err) {
    console.error('password-reset/verify:', err);
    return res.status(500).json({ success: false, message: 'Erro ao verificar código.', error: err.message });
  }
});

app.post('/api/password-reset/request', async (req, res) => {
  try {
    const email = normalizarEmail(req.body?.email);
    if (!email) return res.status(400).json({ success: false, message: 'Email é obrigatório.' });

    const [u] = await pool.query(
      'SELECT ID, EMAIL, NOME FROM SF_USUARIO WHERE EMAIL = ? LIMIT 1',
      [email]
    );

    if (!u.length) return res.status(404).json({ success: false, message: 'Email não cadastrado.' });

    const code = String(Math.floor(100000 + Math.random() * 900000));
    const codeHash = await bcrypt.hash(code, 10);
    const expiresMinutes = 10;

    await pool.query(
      `UPDATE SF_PASSWORD_RESET SET used = 1
        WHERE email = ? AND used = 0`,
      [email]
    );

    await pool.query(
      `INSERT INTO SF_PASSWORD_RESET (email, code_hash, expires_at, used, attempts, token_hash)
       VALUES (?, ?, DATE_ADD(UTC_TIMESTAMP(), INTERVAL ? MINUTE), 0, 0, NULL)`,
      [email, codeHash, expiresMinutes]
    );

    const uid = `reset-${crypto.randomBytes(12).toString('hex')}@sociedadefranciosi`;

    await pool.query(
      `INSERT INTO SF_EMAIL_QUEUE
        (tipo, status, tentativas, max_tentativas,
         id_agendamento, id_usuario, email, nome,
         sala, inicio, fim, motivo, uid, sequence, created_at)
       VALUES
        ('RESET_SENHA', 'PENDENTE', 0, 5,
         NULL, NULL, ?, ?,
         NULL, NULL, NULL, ?, ?, 0, UTC_TIMESTAMP())`,
      [
        email,
        u[0].NOME || email,
        `Seu código de redefinição de senha é: ${code} (expira em ${expiresMinutes} min)`,
        uid
      ]
    );

    return res.json({ success: true });
  } catch (err) {
    console.error('password-reset/request:', err);
    return res.status(500).json({ success: false, message: 'Erro ao solicitar redefinição.', error: err.message });
  }
});

// =====================
// MARKETING (Volume /publicidade)
// =====================
// Volume montado em /publicidade (conforme seu Railway)
const DIRETORIO_VOLUME_PUBLICIDADE = process.env.RAILWAY_VOLUME_MOUNT_PATH || "/publicidade";
const PASTA_MARKETING = path.join(DIRETORIO_VOLUME_PUBLICIDADE, "marketing");

fs.mkdirSync(PASTA_MARKETING, { recursive: true });

// Servir imagens via URL (Express static) [web:650]
app.use("/publicidade/marketing", express.static(PASTA_MARKETING));

function apenasNomeArquivoSeguro(nome) {
  const base = path.basename(String(nome || ""));
  return base.replace(/[^\w.\-() ]+/g, "_");
}
function ehImagem(mimetype) {
  return typeof mimetype === "string" && mimetype.startsWith("image/");
}

const storage = multer.diskStorage({
  destination: (req, file, cb) => cb(null, PASTA_MARKETING),
  filename: (req, file, cb) => {
    const original = apenasNomeArquivoSeguro(file.originalname || "imagem");
    const ext = path.extname(original);
    const nomeSemExt = path.basename(original, ext);
    const unico = `${Date.now()}-${Math.random().toString(16).slice(2)}`;
    cb(null, `${nomeSemExt}-${unico}${ext}`);
  },
});

const upload = multer({
  storage,
  limits: { fileSize: 10 * 1024 * 1024 },
  fileFilter: (req, file, cb) => {
    if (!ehImagem(file.mimetype)) return cb(new Error("Apenas imagens são permitidas."));
    cb(null, true);
  },
});

// LISTAR
app.get("/api/marketing/imagens", async (req, res) => {
  try {
    const files = await fs.promises.readdir(PASTA_MARKETING, { withFileTypes: true });

    const items = files
      .filter((d) => d.isFile())
      .map((d) => d.name)
      .filter((n) => /\.(png|jpe?g|gif|webp|bmp|svg)$/i.test(n))
      .sort((a, b) => a.localeCompare(b, "pt-BR"))
      .map((name) => ({
        name,
        url: `/publicidade/marketing/${encodeURIComponent(name)}`,
      }));

    return res.json({ success: true, items });
  } catch (err) {
    return res.status(500).json({ success: false, message: "Erro ao listar imagens.", error: err.message });
  }
});

// UPLOAD (múltiplos) - campo FormData: "files" [web:647]
app.post("/api/marketing/imagens", upload.array("files", 20), async (req, res) => {
  try {
    const arquivos = Array.isArray(req.files) ? req.files : [];

    const items = arquivos.map((f) => ({
      name: f.filename,
      url: `/publicidade/marketing/${encodeURIComponent(f.filename)}`,
      size: f.size,
      mimetype: f.mimetype,
    }));

    return res.status(201).json({ success: true, items });
  } catch (err) {
    return res.status(400).json({ success: false, message: err.message || "Erro ao enviar imagens." });
  }
});

// REMOVER
app.delete("/api/marketing/imagens/:nome", async (req, res) => {
  try {
    const nome = apenasNomeArquivoSeguro(req.params.nome);
    if (!nome) return res.status(400).json({ success: false, message: "Nome inválido." });

    const base = path.resolve(PASTA_MARKETING);
    const alvo = path.resolve(path.join(PASTA_MARKETING, nome));
    if (!alvo.startsWith(base + path.sep)) {
      return res.status(400).json({ success: false, message: "Caminho inválido." });
    }

    await fs.promises.unlink(alvo);
    return res.json({ success: true, message: "Imagem removida." });
  } catch (err) {
    if (err.code === "ENOENT") {
      return res.status(404).json({ success: false, message: "Arquivo não encontrado." });
    }
    return res.status(500).json({ success: false, message: "Erro ao remover imagem.", error: err.message });
  }
});

function normalizarUF(uf) {
  const s = (uf || '').toString().trim().toUpperCase();
  return s.length === 2 ? s : '';
}

function normalizarDocumento(doc) {
  return (doc || '').toString().replace(/\D+/g, '').trim(); // só números
}

function str(v) {
  const s = (v ?? '').toString().trim();
  return s ? s : '';
}

// GET /api/clientes?q=texto
app.get('/api/clientes', async (req, res) => {
  try {
    const q = (req.query.q || '').toString().trim();

    let sql = `
      SELECT
        ID, RAZAO_SOCIAL, DOCUMENTO, GRUPO_ECONOMICO,
        CIDADE, UF,
        CONTATO_NOME, CONTATO_TELEFONE, CONTATO_EMAIL,
        CULTURA_PRINCIPAL, HECTARES_ESTIMADOS, OBSERVACOES,
        ACTIVE, CREATED_AT, UPDATED_AT
      FROM SF_CLIENTE
      WHERE ACTIVE = 1
    `;
    const params = [];

    if (q) {
      sql += ` AND (RAZAO_SOCIAL LIKE ? OR DOCUMENTO LIKE ?) `;
      params.push(`%${q}%`, `%${normalizarDocumento(q)}%`);
    }

    sql += ` ORDER BY RAZAO_SOCIAL ASC `;

    const [rows] = await pool.query(sql, params);
    return res.json({ success: true, items: rows });
  } catch (err) {
    return res.status(500).json({ success: false, message: 'Erro ao listar clientes.', error: err.message });
  }
});

// GET /api/clientes/:id
app.get('/api/clientes/:id(\\d+)', async (req, res) => {
  try {
    const id = Number(req.params.id);
    const [rows] = await pool.query(
      `SELECT * FROM SF_CLIENTE WHERE ID = ? LIMIT 1`,
      [id]
    );
    if (!rows.length) return res.status(404).json({ success: false, message: 'Cliente não encontrado.' });
    return res.json({ success: true, item: rows[0] });
  } catch (err) {
    return res.status(500).json({ success: false, message: 'Erro ao buscar cliente.', error: err.message });
  }
});


app.post('/api/clientes/salvar', async (req, res) => {
  const conn = await pool.getConnection();
  try {
    const c = req.body?.cliente || {};
    const filiais = Array.isArray(req.body?.filiais) ? req.body.filiais : [];

    const idCliente = Number(c.id || 0) || null;

    const razao = str(c.razao_social);
    const documento = normalizarDocumento(c.documento); // remove máscara (cpf/cnpj)
    const grupo = str(c.grupo_economico) || null;

    const cidade = str(c.cidade);
    const uf = normalizarUF(c.uf);

    const contatoNome = str(c.contato_nome) || null;
    const contatoTelefone = str(c.contato_telefone) || null;
    const contatoEmail = str(c.contato_email) || null;

    const cultura = str(c.cultura_principal) || null;
    const hectaresNum = Number(c.hectares_estimados);
    const hectares = Number.isFinite(hectaresNum) ? hectaresNum : null;
    const obs = str(c.observacoes) || null;

    if (!razao) return res.status(400).json({ success: false, message: 'razao_social é obrigatório.' });
    if (!documento) return res.status(400).json({ success: false, message: 'documento é obrigatório.' });
    if (!cidade) return res.status(400).json({ success: false, message: 'cidade é obrigatória.' });
    if (!uf) return res.status(400).json({ success: false, message: 'uf inválida (2 letras).' });

    await conn.beginTransaction();

    let idFinal = idCliente;

    if (idFinal) {
      const [r] = await conn.query(
        `UPDATE SF_CLIENTE
            SET RAZAO_SOCIAL = ?, DOCUMENTO = ?, GRUPO_ECONOMICO = ?,
                CIDADE = ?, UF = ?,
                CONTATO_NOME = ?, CONTATO_TELEFONE = ?, CONTATO_EMAIL = ?,
                CULTURA_PRINCIPAL = ?, HECTARES_ESTIMADOS = ?, OBSERVACOES = ?
          WHERE ID = ?`,
        [
          razao, documento, grupo,
          cidade, uf,
          contatoNome, contatoTelefone, contatoEmail,
          cultura, hectares, obs,
          idFinal
        ]
      );

      if (r.affectedRows === 0) {
        await conn.rollback();
        return res.status(404).json({ success: false, message: 'Cliente não encontrado.' });
      }
    } else {
      const [r] = await conn.query(
        `INSERT INTO SF_CLIENTE
         (RAZAO_SOCIAL, DOCUMENTO, GRUPO_ECONOMICO, CIDADE, UF,
          CONTATO_NOME, CONTATO_TELEFONE, CONTATO_EMAIL,
          CULTURA_PRINCIPAL, HECTARES_ESTIMADOS, OBSERVACOES, ACTIVE)
         VALUES (?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, 1)`,
        [
          razao, documento, grupo, cidade, uf,
          contatoNome, contatoTelefone, contatoEmail,
          cultura, hectares, obs
        ]
      );
      idFinal = r.insertId;
    }

    // ---------- sincronizar filiais ----------
    const [exist] = await conn.query(
      `SELECT ID FROM SF_CLIENTE_FILIAL WHERE ID_CLIENTE = ? AND ACTIVE = 1`,
      [idFinal]
    );

    const idsExistentes = exist.map(x => Number(x.ID)).filter(n => Number.isFinite(n) && n > 0);
    const idsFormulario = filiais.map(f => Number(f.id || 0)).filter(n => Number.isFinite(n) && n > 0);

    const idsParaDesativar = idsExistentes.filter(id => !idsFormulario.includes(id));
    if (idsParaDesativar.length) {
      await conn.query(
        `UPDATE SF_CLIENTE_FILIAL SET ACTIVE = 0 WHERE ID_CLIENTE = ? AND ID IN (?)`,
        [idFinal, idsParaDesativar]
      );
    }

    for (const f of filiais) {
      const fid = Number(f.id || 0) || null;

      const nome = str(f.nome);
      const endereco = str(f.endereco) || null;
      const fCidade = str(f.cidade);
      const fUf = normalizarUF(f.uf);
      const fContatoNome = str(f.contato_nome) || null;
      const fContatoTelefone = str(f.contato_telefone) || null;

      if (!nome) { await conn.rollback(); return res.status(400).json({ success: false, message: 'Filial: nome é obrigatório.' }); }
      if (!fCidade) { await conn.rollback(); return res.status(400).json({ success: false, message: 'Filial: cidade é obrigatória.' }); }
      if (!fUf) { await conn.rollback(); return res.status(400).json({ success: false, message: 'Filial: uf inválida (2 letras).' }); }

      if (fid) {
        await conn.query(
          `UPDATE SF_CLIENTE_FILIAL
              SET NOME = ?, ENDERECO = ?, CIDADE = ?, UF = ?,
                  CONTATO_NOME = ?, CONTATO_TELEFONE = ?, ACTIVE = 1
            WHERE ID = ? AND ID_CLIENTE = ?`,
          [nome, endereco, fCidade, fUf, fContatoNome, fContatoTelefone, fid, idFinal]
        );
      } else {
        await conn.query(
          `INSERT INTO SF_CLIENTE_FILIAL
           (ID_CLIENTE, NOME, ENDERECO, CIDADE, UF, CONTATO_NOME, CONTATO_TELEFONE, ACTIVE)
           VALUES (?, ?, ?, ?, ?, ?, ?, 1)`,
          [idFinal, nome, endereco, fCidade, fUf, fContatoNome, fContatoTelefone]
        );
      }
    }

    await conn.commit();
    return res.status(200).json({ success: true, id: idFinal });
  } catch (err) {
    try { await conn.rollback(); } catch {}

    if (err.code === 'ER_DUP_ENTRY') {
      return res.status(409).json({ success: false, message: 'Já existe cliente com este documento.' });
    }

    return res.status(500).json({ success: false, message: 'Erro ao salvar cliente.', error: err.message });
  } finally {
    conn.release();
  }
});

// PUT /api/clientes/:id
app.put('/api/clientes/:id(\\d+)', async (req, res) => {
  try {
    const id = Number(req.params.id);

    const razao = str(req.body?.razao_social);
    const documento = normalizarDocumento(req.body?.documento);
    const grupo = str(req.body?.grupo_economico) || null;

    const cidade = str(req.body?.cidade);
    const uf = normalizarUF(req.body?.uf);

    const contatoNome = str(req.body?.contato_nome) || null;
    const contatoTelefone = str(req.body?.contato_telefone) || null;
    const contatoEmail = str(req.body?.contato_email) || null;

    const cultura = str(req.body?.cultura_principal) || null;
    const hectares = Number(req.body?.hectares_estimados);
    const obs = str(req.body?.observacoes) || null;

    if (!razao) return res.status(400).json({ success: false, message: 'razao_social é obrigatório.' });
    if (!documento) return res.status(400).json({ success: false, message: 'documento é obrigatório.' });
    if (!cidade) return res.status(400).json({ success: false, message: 'cidade é obrigatória.' });
    if (!uf) return res.status(400).json({ success: false, message: 'uf inválida (2 letras).' });

    const [r] = await pool.query(
      `UPDATE SF_CLIENTE
          SET RAZAO_SOCIAL = ?, DOCUMENTO = ?, GRUPO_ECONOMICO = ?,
              CIDADE = ?, UF = ?,
              CONTATO_NOME = ?, CONTATO_TELEFONE = ?, CONTATO_EMAIL = ?,
              CULTURA_PRINCIPAL = ?, HECTARES_ESTIMADOS = ?, OBSERVACOES = ?
        WHERE ID = ?`,
      [
        razao, documento, grupo,
        cidade, uf,
        contatoNome, contatoTelefone, contatoEmail,
        cultura, Number.isFinite(hectares) ? hectares : null, obs,
        id
      ]
    );

    if (r.affectedRows === 0) return res.status(404).json({ success: false, message: 'Cliente não encontrado.' });
    return res.json({ success: true });
  } catch (err) {
    if (err.code === 'ER_DUP_ENTRY') {
      return res.status(409).json({ success: false, message: 'Já existe cliente com este documento.' });
    }
    return res.status(500).json({ success: false, message: 'Erro ao atualizar cliente.', error: err.message });
  }
});

// DELETE /api/clientes/:id
app.delete('/api/clientes/:id(\\d+)', async (req, res) => {
  const conn = await pool.getConnection();
  try {
    const id = Number(req.params.id);

    await conn.beginTransaction();

    // desativa filiais ativas
    await conn.query(
      `UPDATE SF_CLIENTE_FILIAL SET ACTIVE = 0 WHERE ID_CLIENTE = ? AND ACTIVE = 1`,
      [id]
    );

    // desativa cliente
    const [r] = await conn.query(
      `UPDATE SF_CLIENTE SET ACTIVE = 0 WHERE ID = ? AND ACTIVE = 1`,
      [id]
    );

    if (r.affectedRows === 0) {
      await conn.rollback();
      return res.status(404).json({ success: false, message: 'Cliente não encontrado ou já desativado.' });
    }

    await conn.commit();
    return res.json({ success: true });
  } catch (err) {
    try { await conn.rollback(); } catch {}
    return res.status(500).json({ success: false, message: 'Erro ao desativar cliente.', error: err.message });
  } finally {
    conn.release();
  }
});

// GET /api/clientes/:id/filiais
app.get('/api/clientes/:id(\\d+)/filiais', async (req, res) => {
  try {
    const idCliente = Number(req.params.id);

    const [rows] = await pool.query(
      `SELECT ID, ID_CLIENTE, NOME, ENDERECO, CIDADE, UF, CONTATO_NOME, CONTATO_TELEFONE, ACTIVE, CREATED_AT, UPDATED_AT
         FROM SF_CLIENTE_FILIAL
        WHERE ID_CLIENTE = ? AND ACTIVE = 1
        ORDER BY NOME ASC`,
      [idCliente]
    );

    return res.json({ success: true, items: rows });
  } catch (err) {
    return res.status(500).json({ success: false, message: 'Erro ao listar filiais.', error: err.message });
  }
});

// POST /api/clientes/:id/filiais
app.post('/api/clientes/:id(\\d+)/filiais', async (req, res) => {
  try {
    const idCliente = Number(req.params.id);

    const nome = str(req.body?.nome);
    const endereco = str(req.body?.endereco) || null;
    const cidade = str(req.body?.cidade);
    const uf = normalizarUF(req.body?.uf);
    const contatoNome = str(req.body?.contato_nome) || null;
    const contatoTelefone = str(req.body?.contato_telefone) || null;

    if (!nome) return res.status(400).json({ success: false, message: 'nome é obrigatório.' });
    if (!cidade) return res.status(400).json({ success: false, message: 'cidade é obrigatória.' });
    if (!uf) return res.status(400).json({ success: false, message: 'uf inválida (2 letras).' });

    const [r] = await pool.query(
      `INSERT INTO SF_CLIENTE_FILIAL
       (ID_CLIENTE, NOME, ENDERECO, CIDADE, UF, CONTATO_NOME, CONTATO_TELEFONE, ACTIVE)
       VALUES (?, ?, ?, ?, ?, ?, ?, 1)`,
      [idCliente, nome, endereco, cidade, uf, contatoNome, contatoTelefone]
    );

    return res.status(201).json({ success: true, item: { id: r.insertId } });
  } catch (err) {
    return res.status(500).json({ success: false, message: 'Erro ao criar filial.', error: err.message });
  }
});

// PUT /api/filiais/:id
app.put('/api/filiais/:id(\\d+)', async (req, res) => {
  try {
    const id = Number(req.params.id);

    const nome = str(req.body?.nome);
    const endereco = str(req.body?.endereco) || null;
    const cidade = str(req.body?.cidade);
    const uf = normalizarUF(req.body?.uf);
    const contatoNome = str(req.body?.contato_nome) || null;
    const contatoTelefone = str(req.body?.contato_telefone) || null;

    if (!nome) return res.status(400).json({ success: false, message: 'nome é obrigatório.' });
    if (!cidade) return res.status(400).json({ success: false, message: 'cidade é obrigatória.' });
    if (!uf) return res.status(400).json({ success: false, message: 'uf inválida (2 letras).' });

    const [r] = await pool.query(
      `UPDATE SF_CLIENTE_FILIAL
          SET NOME = ?, ENDERECO = ?, CIDADE = ?, UF = ?, CONTATO_NOME = ?, CONTATO_TELEFONE = ?
        WHERE ID = ?`,
      [nome, endereco, cidade, uf, contatoNome, contatoTelefone, id]
    );

    if (r.affectedRows === 0) return res.status(404).json({ success: false, message: 'Filial não encontrada.' });
    return res.json({ success: true });
  } catch (err) {
    return res.status(500).json({ success: false, message: 'Erro ao atualizar filial.', error: err.message });
  }
});

// DELETE /api/filiais/:id
app.delete('/api/filiais/:id(\\d+)', async (req, res) => {
  try {
    const id = Number(req.params.id);
    const [r] = await pool.query(`UPDATE SF_CLIENTE_FILIAL SET ACTIVE = 0 WHERE ID = ? AND ACTIVE = 1`, [id]);
    if (r.affectedRows === 0) return res.status(404).json({ success: false, message: 'Filial não encontrada ou já desativada.' });
    return res.json({ success: true });
  } catch (err) {
    return res.status(500).json({ success: false, message: 'Erro ao desativar filial.', error: err.message });
  }
});




// =====================
// Inicia servidor (sempre por último)
// =====================
app.listen(PORT, () => {
  console.log(`🚀 API rodando na porta ${PORT}`);
  console.log('✅ Teste: /health');
});
