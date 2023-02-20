-- This SQLite script is used to create the relevant tables for logging user information 
-- and usage of the dExcel add-in as well as the dExcel installer.
CREATE TABLE IF NOT EXISTS users(
    username NVARCHAR(255) PRIMARY KEY,
    firstname NVARCHAR(255),
    surname NVARCHAR(255),
    date_created NVARCHAR(255),
    active BOOLEAN
);

CREATE TABLE IF NOT EXISTS dexcel_usage(
    username NVARCHAR(255),
    version NVARCHAR(255), 
    date_logged NVARCHAR(255)
);

CREATE TABLE IF NOT EXISTS dexcel_installer_usage(
    username NVARCHAR(255),
    version NVARCHAR(255), 
    date_logged NVARCHAR(255)
);
