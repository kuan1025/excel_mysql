-- create database EXCEL_TEST;
 USE EXCEL_TEST;
 select * from test1;
create table test1(
id int auto_increment primary key,
username varchar(10) not null,
age int not null,
pc blob not null
);
 drop table test1;
insert into test1 ( `username`,age ) values (  'hong' ,  32  );
