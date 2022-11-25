-- select * from users;


-- INSERT INTO users 
-- VALUES
--     ('monkey')
-- WHERE IF NOT EXISTS (select * from users where username = 'monkey'); create table if not exists users(
create table users(
    username NVARCHAR(255) PRIMARY KEY,
    firstname NVARCHAR(255),
    surname NVARCHAR(255),
    created NVARCHAR(255),
    active BOOLEAN
);

CREATE TABLE dexcel_usage(
    username NVARCHAR(255),
    version NVARCHAR(255), 
    date NVARCHAR(255)
);

CREATE TABLE dexcel_installer_usage(
    username NVARCHAR(255) 

DROP TABLE users;

insert into users(username, firstname, surname, created, active)
select 'stcollins', 'Storm', 'Collins', DATETIME('NOW'), TRUE
WHERE NOT EXISTS (select * from users where username = 'stcollins');

select * from users;

Delete from users where username = 'stcollins';