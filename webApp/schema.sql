drop table if exists projects;

create table projects (
  sheetId integer not null,
  projectId text not null,
  country text not null,
  capexCicle text not null,
  projectManager text not null,
  projectName text not null
);
