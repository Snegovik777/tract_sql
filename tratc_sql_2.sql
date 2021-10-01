SET DATEFORMAT dmy;
select ph.Name as 'Название фонограммы', ph_ref.val_str as 'Автор музыки', 
		ph_ref_2.val_str as 'Автор слов', COUNT('hist.ph_id') as 'Кол-во сообщений в эфир', 
		ph_ref_3.val_str as 'Исполнитель (ФИО исполнителя или название коллектива)', ph_ref_4.val_str as 'Изготовитель фонограммы' from dbo.PH_PLAY_HISTORY hist
left join dbo.ph as ph -- джойн из осн таблицы dbo.PH
on hist.ph_id = ph.id
left join (select ph_id, val_str from dbo.ph_val_reflection where name = 'Автор музыки') as ph_ref
on hist.ph_id = ph_ref.ph_id
left join (select ph_id, val_str from dbo.ph_val_reflection where name = 'Автор слов') as ph_ref_2
on hist.ph_id = ph_ref_2.ph_id
left join (select ph_id, val_str from dbo.ph_val_reflection where name = 'Исполнитель') as ph_ref_3
on hist.ph_id = ph_ref_3.ph_id
left join (select ph_id, val_str from dbo.ph_val_reflection where name = 'Изготовитель фонограммы') as ph_ref_4
on hist.ph_id = ph_ref_4.ph_id
WHERE hist.BlockId like 'PLAYLIST_RM%' and hist.PlayTime between '01.07.2021' and '15.07.2021' and ph_ref.val_str is not null
group by hist.PlayTime, ph.Name, hist.ph_id, ph_ref.val_str, ph_ref_2.val_str, ph_ref_3.val_str, ph_ref_4.val_str
order by hist.PlayTime  -- остсортированы по даты по возрастанию
