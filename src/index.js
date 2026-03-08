app.get('/api/estoque/controle/escritorio', async (req, res) => {
  let conn;

  try {
    conn = await pool.getConnection();

    const [rows] = await conn.query(
      `
      SELECT
        pe.produto_sistema_id AS id,
        COALESCE(p.codigo, pe.cod_produto_sistema) AS codigo_item,
        COALESCE(p.descricao, pe.descricao_produto_nf) AS descricao_item,
        SUM(COALESCE(pe.qtd_nf, 0)) AS qtd_disponivel,
        0 AS qtd_em_pedido,
        pe.LOCAL AS local,
        pe.ID_LOCAL_ALMOXARIFADO AS id_local_almoxarifado
      FROM SF_PRODUTO_ENTRADA pe
      LEFT JOIN SF_PRODUTOS p
        ON p.id = pe.produto_sistema_id
      WHERE UPPER(CONVERT(COALESCE(pe.LOCAL, '') USING utf8mb4)) COLLATE utf8mb4_general_ci LIKE '%ESCRITORIO%'
         OR UPPER(CONVERT(COALESCE(pe.LOCAL, '') USING utf8mb4)) COLLATE utf8mb4_general_ci LIKE '%ESCRITÓRIO%'
      GROUP BY
        pe.produto_sistema_id,
        COALESCE(p.codigo, pe.cod_produto_sistema),
        COALESCE(p.descricao, pe.descricao_produto_nf),
        pe.LOCAL,
        pe.ID_LOCAL_ALMOXARIFADO
      ORDER BY COALESCE(p.codigo, pe.cod_produto_sistema) ASC
      `
    );

    return res.json({
      success: true,
      items: rows
    });
  } catch (err) {
    console.error('Erro ao carregar estoque do escritório:', err);

    return res.status(500).json({
      success: false,
      message: 'Erro ao carregar estoque do escritório.',
      error: err.message
    });
  } finally {
    if (conn) conn.release();
  }
});
