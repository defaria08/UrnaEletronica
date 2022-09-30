--Database: UrnaEletronica

--DROP DATABASE IF EXISTS "UrnaEletronica";

CREATE DATABASE "UrnaEletronica"
    WITH
    OWNER = postgres
    ENCODING = 'UTF8'
    LC_COLLATE = 'Portuguese_Brazil.1252'
    LC_CTYPE = 'Portuguese_Brazil.1252'
    TABLESPACE = pg_default
    CONNECTION LIMIT = -1
    IS_TEMPLATE = False;

create table candidato (
id int not null primary key,
cargo varchar(20) not null,
nome varchar(30) not null,
partido varchar(30) not null,
foto varchar(30) not null);

create table voto (
id serial not null primary key,
deputadoe int,
deputadof int,
governador int,
presidente int);

insert into candidato (id, cargo,nome,partido,foto) values (14123,'DEPUTADO FEDERAL','CARLOS MASSA "RATINHO"','PARTIDO SBT','ratinho.png');
insert into candidato (id, cargo,nome,partido,foto) values (1411,'DEPUTADO ESTADUAL','ELIANA BEZERRA','PARTIDO SBT','eliana.png');
insert into candidato (id, cargo,nome,partido,foto) values (14,'GOVERNADOR','CELSO PORTIOLLI','PARTIDO SBT','portioli.png');
insert into candidato (id, cargo,nome,partido,foto) values (15,'PRESIDENTE','SILVIO SANTOS','PARTIDO SBT','ssantos.png');

insert into candidato (id, cargo,nome,partido,foto) values (17171,'DEPUTADO FEDERAL','TADEU SCHMIDT','PARTIDO GLOBO','tadeu.png');
insert into candidato (id, cargo,nome,partido,foto) values (1717,'DEPUTADO ESTADUAL','GLORIA MARIA','PARTIDO GLOBO','gloria.png');
insert into candidato (id, cargo,nome,partido,foto) values (17,'GOVERNADOR','WILLIAN BONNER','PARTIDO GLOBO','boner.png');
insert into candidato (id, cargo,nome,partido,foto) values (18,'PRESIDENTE','LUCIANO HUCK','PARTIDO GLOBO','luciano.png');

insert into candidato (id, cargo,nome,partido,foto) values (23231,'DEPUTADO FEDERAL','MILTON NEVES','PARTIDO BAND','milton.png');
insert into candidato (id, cargo,nome,partido,foto) values (2323,'DEPUTADO ESTADUAL','RENATA FAN','PARTIDO BAND','renata.png');
insert into candidato (id, cargo,nome,partido,foto) values (23,'GOVERNADOR','JOSE LUIZ DATENA','PARTIDO BAND','datena.png');
insert into candidato (id, cargo,nome,partido,foto) values (22,'PRESIDENTE','FAUSTO SILVA','PARTIDO BAND','fasto.png');