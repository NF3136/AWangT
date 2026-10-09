-- 仅 SELECT，不修改数据库。分别执行下面两段查询。
-- 导出选择“查询数据库获取全部结果”，不要只导出已加载行。
-- 本轮核实已有静态路由是否会被后续号码处理改变；这些表不是用户注册状态。

-- 第一段：导出为 主叫与号码处理配置.csv。
-- CNACLR/PFXPRO/DNC 取同源全部记录，避免漏掉通配规则和间接引用。
-- CNACLD 仅补充上轮未查询的通配号首集 65534。
WITH chosen_source AS (
    SELECT 'JSL_VOBB_AGCFOMU/ALLME_20260803065001.txt'::text AS source_file
), table_list(table_name) AS (
    VALUES ('t_agcf_cnaclr'), ('t_agcf_pfxpro'), ('t_agcf_dnc'), ('t_agcf_cnacld')
), configs AS (
    SELECT 't_agcf_cnaclr'::text AS table_name,
           t.source_file, t.id AS record_id, to_jsonb(t) AS config
    FROM public.t_agcf_cnaclr t
    JOIN chosen_source s ON s.source_file = t.source_file
    UNION ALL
    SELECT 't_agcf_pfxpro', t.source_file, t.id, to_jsonb(t)
    FROM public.t_agcf_pfxpro t
    JOIN chosen_source s ON s.source_file = t.source_file
    UNION ALL
    SELECT 't_agcf_dnc', t.source_file, t.id, to_jsonb(t)
    FROM public.t_agcf_dnc t
    JOIN chosen_source s ON s.source_file = t.source_file
    UNION ALL
    SELECT 't_agcf_cnacld', t.source_file, t.id, to_jsonb(t)
    FROM public.t_agcf_cnacld t
    JOIN chosen_source s ON s.source_file = t.source_file
    WHERE btrim(t.lp) = '65534'
), output AS (
    SELECT 'summary'::text AS record_kind,
           l.table_name, s.source_file, NULL::bigint AS record_id,
           jsonb_build_object('row_count', count(c.record_id)) AS config
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

-- 第二段：导出为 主叫与号码处理参数说明.csv。
SELECT command_name, "参数标识", "参数名称", "参数说明"
FROM public.command_info_uac3000
WHERE upper(btrim(command_name)) ~ '(^|[[:space:]])(CNACLR|PFXPRO|DNC)$'
ORDER BY command_name, "参数标识", id;
