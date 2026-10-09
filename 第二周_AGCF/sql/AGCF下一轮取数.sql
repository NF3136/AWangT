-- 用途：核对同一来源的静态路由与链路配置。所有语句只读。
-- 分别选中第一段、第二段执行，各导出一个 CSV；不要把整个文件当作一条语句执行。
-- 这些数据不能独立证明主叫当前注册的 S-CSCF，也不直接提供 PAI/PPK。

-- 第一段：保存为「同源链路配置.csv」。
-- 一次取指定 source_file 的六张配置表，保留来源和全部已有字段。
-- SIPCFG 已由「各表样例数据.csv」提供，本轮不重复取数。
-- 不取 ASBR 的认证信息，不跨站点拼接记录，不截断为前 50 条。
-- summary 行会给出每张表在此 source_file 下的行数；零行也会显示。
WITH chosen_source AS (
    SELECT 'JSL_VOBB_AGCFOMU/ALLME_20260803065001.txt'::text AS source_file
), table_list(table_name) AS (
    VALUES ('t_agcf_callsrc'), ('t_agcf_rt'), ('t_agcf_rtana'),
           ('t_agcf_srt'), ('t_agcf_siptg'), ('t_agcf_sipippair')
), configs AS (
    SELECT 't_agcf_callsrc' AS table_name, t.source_file,
           t.id AS record_id, to_jsonb(t) AS config
    FROM public.t_agcf_callsrc AS t
    JOIN chosen_source AS s ON t.source_file = s.source_file
    UNION ALL
    SELECT 't_agcf_rt', t.source_file, t.id, to_jsonb(t)
    FROM public.t_agcf_rt AS t
    JOIN chosen_source AS s ON t.source_file = s.source_file
    UNION ALL
    SELECT 't_agcf_rtana', t.source_file, t.id, to_jsonb(t)
    FROM public.t_agcf_rtana AS t
    JOIN chosen_source AS s ON t.source_file = s.source_file
    UNION ALL
    SELECT 't_agcf_srt', t.source_file, t.id, to_jsonb(t)
    FROM public.t_agcf_srt AS t
    JOIN chosen_source AS s ON t.source_file = s.source_file
    UNION ALL
    SELECT 't_agcf_siptg', t.source_file, t.id, to_jsonb(t)
    FROM public.t_agcf_siptg AS t
    JOIN chosen_source AS s ON t.source_file = s.source_file
    UNION ALL
    SELECT 't_agcf_sipippair', t.source_file, t.id, to_jsonb(t)
    FROM public.t_agcf_sipippair AS t
    JOIN chosen_source AS s ON t.source_file = s.source_file
), output AS (
    SELECT 'summary'::text AS record_kind, l.table_name,
           s.source_file, NULL::bigint AS record_id,
           jsonb_build_object('row_count', count(c.record_id)) AS config
    FROM table_list AS l
    CROSS JOIN chosen_source AS s
    LEFT JOIN configs AS c ON c.table_name = l.table_name
    GROUP BY l.table_name, s.source_file
    UNION ALL
    SELECT 'config', table_name, source_file, record_id, config
    FROM configs
)
SELECT record_kind, table_name, source_file, record_id, config
FROM output
ORDER BY table_name, CASE WHEN record_kind = 'summary' THEN 0 ELSE 1 END,
         record_id;


-- 第二段：保存为「链路参数说明.csv」。
-- 先使用与 AGCF/UAC 对应的说明，不将 CSC3300/BCF 的同名命令直接套用。
-- ATKDNSEG 是此前 TGAP 参数说明引用的中继号段命令，顺带核对是否有说明。
SELECT command_name, "参数标识", "参数名称", "参数说明"
FROM public.command_info_uac3000
WHERE UPPER(TRIM(command_name)) ~
      '(^|[[:space:]])(CALLSRC|RTANA|RT|SRT|SIPTG|SIPIPPAIR|SIPCFG|ATKDNSEG)$'
ORDER BY command_name, "参数标识", id;
