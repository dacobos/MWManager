drop table if exists projects;
drop table if exists users;

create table projects (
  sheetId integer not null,
  projectId text not null,
  country text not null,
  capexCicle text not null,
  projectManager text not null,
  projectName text not null
);

create table users (
  name text not null,
  lastName text not null,
  role text not null,
  email text not null unique,
  password text,
  question text,
  secret text,
  token text
);
