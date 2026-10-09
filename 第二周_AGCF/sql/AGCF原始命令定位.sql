-- 仅查询表结构，不修改数据库，也不读取认证信息。
-- 导出为 原始命令表结构.csv。
-- agcf_order 已出现在当前数据库的表列表中，但尚未确认其内容。
-- 此处不能仅凭表名假定它保存了原始 MML 命令。
SELECT table_schema, table_name, ordinal_position, column_name, data_type
FROM information_schema.columns
WHERE table_schema = 'public'
  AND table_name = 'agcf_order'
ORDER BY ordinal_position;
