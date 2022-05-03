Create Database WBExchange
use WBExchange
Go

create table WB_Exchange_MessageTrackingLog(

ID int primary key identity NOT NULL,
MTLog_Timestamp datetime NOT NULL,
MTLog_Sender nvarchar(max) NOT NULL,
MTLog_Recipients nvarchar(max) NOT NULL,
MTLog_MsgSubject nvarchar(max) NOT NULL,
MTLog_Source nvarchar(max) NOT NULL,
MTLog_EventID nvarchar(max) NOT NULL,
MTLog_SrcCntxt nvarchar(max) NULL,
MTLog_MsgId nvarchar(max) NULL,
MTLog_Save_Time DateTime default getdate() null

);


select * from WB_Exchange_MessageTrackingLog

select * from WB_Exchange_MessageTrackingLog where MTLog_Sender='richy@wb-coatings.de' or MTLog_Recipients='richy@wb-coatings.de'

delete from WB_Exchange_MessageTrackingLog

select distinct MTLog_EventID from WB_Exchange_MessageTrackingLog

select * from WB_Exchange_MessageTrackingLog where MTLog_EventID='DEliver'

select * from WB_Exchange_MessageTrackingLog where MTLog_Recipients like 'ri%'

select * from WB_Exchange_MessageTrackingLog where MTLog_Timestamp >= '2022/03/13 00:00:00'  AND MTLog_Timestamp <= '2022/03/13 23:59:59' order by MTLog_Timestamp desc

select * from WB_Exchange_MessageTrackingLog where MTLog_Timestamp >= '2022/03/10 00:00:00'  AND MTLog_Timestamp <= '2022/03/13 23:59:59'

CREATE PROCEDURE MsgTrackingLog AS select * from WB_Exchange_MessageTrackingLog where MTLog_Sender='richy@wb-coatings.de' or MTLog_Recipients='richy@wb-coatings.de'

EXEC MsgTrackingLog


select * from WB_Exchange_MessageTrackingLog where MTLog_Timestamp >= '2022-03-15 00:00:00'

select * from WB_Exchange_MessageTrackingLog

select * from WB_Exchange_MessageTrackingLog where MTLog_Sender = 'richy@wb-coatings.de' and MTLog_Recipients ='richy@wb-coatings.de' and MTLog_Timestamp BETWEEN '2022-03-15 00:00:00' AND '2022-03-14 00:00:00'

select * from WB_Exchange_MessageTrackingLog where MTLog_Timestamp >= '2022-03-13 00:00:00' and (MTLog_Sender = 'richy@wb-coatings.de' or MTLog_Recipients ='richy@wb-coatings.de') 

select * from WB_Exchange_MessageTrackingLog where MTLog_MsgSubject Like 'F%'

select * from WB_Exchange_MessageTrackingLog where MTLog_Timestamp BETWEEN '2022-03-19 00:00:00' AND '2022-03-31 23:59:59' order by MTLog_Timestamp desc