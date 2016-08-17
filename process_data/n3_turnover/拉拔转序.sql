select
    *, 'W3下料' as 工序
from
    [成品自主下料$d2:al]
where
    合计 <> 0 and
    产品图号 is not null

union all

select
    *, 'W3T2' as 工序
from
    [一部转出$d2:al]
where
    合计 <> 0 and
    产品图号 is not null

union all

select
    *, 'W3T1' as 工序
from
    [G加转出$d2:al]
where
    合计 <> 0 and
    产品图号 is not null
