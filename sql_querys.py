def get_header(gsalid):
    return f"""
        
			  select g.stype,convert(varchar(100),g.WRKORDNO)+'/'+convert(varchar(100),b.grecno) as wrkordno,b.BILLD,b1.PRERECNO,v.LICNO,v.make+' '+v.model as model,v.SERIALNO,g.CREATED,g.DISTDRIV,b1.PREBILLD,b.BILLD 
			  , isnull(b.prerecno,g.wrkordno) as orderno
			  from GSALS01  g
join GBILS01 b on case when g.stype='u' and b.BTYPE=98 then 1
                                     when g.stype<>'u' then 1
                                     end =1
                                     and b. GSALID=g.GSALID 
join vehi v on v.vehiid=g.VEHIID
left join GBILS01 b1 on b1.PRERECNO=b.PRERECNO and b1.BTYPE=99

where g.GSALID={gsalid}

    """


def central_table(gsalid):
    return f"""
        select g.workid,g.name,1,g.num , sum(g.num) over (partition by g.workid) as vsogo_ng,((sum(g.num) over ()  * 560)) as vartist
from grows01  g
join gsals01 s on s.GSALID=g.GSALID

where g.gsalid={gsalid} and RTYPE in (7,4) 
and case when s.stype='u' and g.btype=98 then 1
              when s.stype<>'u' then 1
              end=1
              
    """


def footer(gsalid):
    return f"""
      select g.itemno,g.name,i.TARIFFNO, 'шт' ,g.num,round(u.f2/1.2,2) as basepr,33.78 as exrate,d.ADINDATA as current_rate, round(g.UNITPR/1.2,2) unitpr,  round(g.rsum/1.2,2)
from grows01  g
join gsals01 s on s.GSALID=g.GSALID
join item i on i.itemno=g.itemno and i.SUPLNO=g.SUPLNO 
left join amintegrations.dbo.UPbaseprice u on u.f1=g.itemno
left join adin d on d.INSTYPE='g' and FIELDID='006' and d.ADINID=g.GSALID 
where g.gsalid={gsalid} 
and case when s.stype='u' and g.btype=98 and g.rtype=2 then 1
              when s.stype<>'u' and  g.rtype=1 then 1
              end=1 """
