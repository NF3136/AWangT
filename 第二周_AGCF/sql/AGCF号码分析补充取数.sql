-- 目的：核实中继 38 → CALLSRC 30 关联的号码分析条件。
-- 仅查询；不修改数据库。本文件有两段独立查询，分别执行、导出。
-- 导出时选择查询数据库获取全部结果，不要只导出已经加载的 200 行。

-- 第一段：导出为 号码分析配置.csv。
-- 保留来源、原始 ADD/MOD 行，并带汇总行，便于核对导出是否完整。
WITH chosen_source AS (
    SELECT 'JSL_VOBB_AGCFOMU/ALLME_20260803065001.txt'::text AS source_file
), table_list(table_name) AS (
    VALUES ('t_agcf_ldnset'), ('t_agcf_cnacld')
), configs AS (
    SELECT 't_agcf_ldnset'::text AS table_name,
           t.source_file, t.id AS record_id, to_jsonb(t) AS config
    FROM public.t_agcf_ldnset t
    JOIN chosen_source s ON s.source_file = t.source_file
    WHERE btrim(t.lp) = '30'
    UNION ALL
    SELECT 't_agcf_cnacld'::text,
           t.source_file, t.id, to_jsonb(t)
    FROM public.t_agcf_cnacld t
    JOIN chosen_source s ON s.source_file = t.source_file
    WHERE btrim(t.lp) = '30'
), output AS (
    SELECT 'summary'::text AS record_kind,
           l.table_name, s.source_file, NULL::bigint AS record_id,
           jsonb_build_object('row_count', count(c.record_id), 'lp_filter', '30') AS config
    FROM table_list l
    CROSS JOIN chosen_source s
    LEFT JOIN configs c
      ON c.table_name = l.table_name AND c.source_file = s.source_file
    GROUP BY l.table_name, s.source_file
    UNION ALL
    SELECT 'config', table_name, source_file, record_id, config
    FROM configs
)
SELECT record_kind, table_name, source_file, record_id, config
FROM output
ORDER BY table_name,
         CASE record_kind WHEN 'summary' THEN 0 ELSE 1 END,
         record_id;

-- 第二段：导出为 号码分析参数说明.csv。
-- 同时获取参数定义，避免仅凭字段缩写推断号码匹配和路由规则。
SELECT command_name, "参数标识", "参数名称", "参数说明"
FROM public.command_info_uac3000
WHERE upper(btrim(command_name)) ~ '(^|[[:space:]])(CNACLD|LDNSET)$'
ORDER BY command_name, "参数标识", id;
