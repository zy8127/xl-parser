select
    产品图号, 上月结存 as 数量, 'W3盘点' as 工序
from
    [汇总$d2:f]
where
    上月结存 <> 0 and
    产品图号 is not null
