
drop database pythonProject
create database pythonProject

use pythonProject
create table myProductsTB
(
name nvarchar(50),
price int
)

select *
from myProductsTB

--delete from myProductsTB