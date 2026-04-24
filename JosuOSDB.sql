USE master;
GO
IF EXISTS (SELECT * FROM sys.databases WHERE name = 'JosuOsDB')
BEGIN
    DROP DATABASE JosuOsDB;
END
GO

CREATE DATABASE JosuOsDB;
GO
USE JosuOsDB;
GO

-- 1. Tabla de Usuarios (Login)
CREATE TABLE Usuarios (
    Id INT PRIMARY KEY IDENTITY(1,1),
    Usuario NVARCHAR(50) UNIQUE NOT NULL,
    Contrasena NVARCHAR(50) NOT NULL,
    Foto VARBINARY(MAX)
);

-- 2. Tabla de Docentes
CREATE TABLE Docentes (
    IdDocente INT PRIMARY KEY IDENTITY(1,1),
    NombreFull NVARCHAR(100) NOT NULL,
    Cedula NVARCHAR(20) UNIQUE NOT NULL,
    Especialidad NVARCHAR(50), 
    Telefono NVARCHAR(15),
    Email NVARCHAR(100)
);

-- 3. Tabla de Materias (Con relación al docente y horario)
CREATE TABLE Materias (
    IdMateria INT PRIMARY KEY IDENTITY(1,1),
    NombreMateria NVARCHAR(100) NOT NULL,
    Creditos INT,
    IdDocente INT, -- Relación con el profesor
    Horario NVARCHAR(50), -- Bloque de la American College
    CONSTRAINT FK_Docente FOREIGN KEY (IdDocente) REFERENCES Docentes(IdDocente)
);

-- 4. Datos iniciales para pruebas
INSERT INTO Usuarios (Usuario, Contrasena) VALUES ('admin', '1234');

-- Opcional: Un docente de prueba para que el ComboBox no esté vacío
INSERT INTO Docentes (NombreFull, Cedula, Especialidad, Telefono, Email) 
VALUES ('Docente de Prueba', '000-000000-0000X', 'Sistemas', '88888888', 'prueba@uac.edu.ni');

INSERT INTO Docentes (NombreFull, Cedula, Especialidad, Telefono, Email) VALUES 
('Juan Carlos Pérez', '001-120585-0001A', 'Ingeniería en Sistemas', '8845-1234', 'jperez@uac.edu.ni'),
('María Auxiliadora López', '001-201090-0005B', 'Derecho', '7756-9876', 'mlopez@uac.edu.ni'),
('Ricardo Antonio Ortega', '041-150878-0002C', 'Administración de Empresas', '8233-4455', 'rortega@uac.edu.ni'),
('Claudia Beatriz Mendoza', '281-050495-1003D', 'Marketing', '5844-3322', 'cmendoza@uac.edu.ni'),
('Franklin José Espinoza', '001-251288-0009E', 'Diseño Gráfico', '8900-1122', 'fespinoza@uac.edu.ni'),
('Irene Martínez Mejía', '001-010175-0000F', 'Ingeniería en Sistemas', '8765-4321', 'imartinez@uac.edu.ni'),
('Roberto Carlos Méndez', '161-120982-0008G', 'Derecho', '7512-3456', 'rmendez@uac.edu.ni'),
('Sonia del Carmen Ruiz', '001-300692-0004H', 'Diseño Gráfico', '8422-5566', 'sruiz@uac.edu.ni');

GO



USE JosuOsDB;
GO

-- 1. Tabla para Control de Asistencia
CREATE TABLE Asistencias (
    IdAsistencia INT PRIMARY KEY IDENTITY(1,1),
    IdMateria INT,
    Fecha DATE DEFAULT GETDATE(),
    Estado NVARCHAR(20), -- 'Presente', 'Falta', 'Justificado'
    Observaciones NVARCHAR(MAX),
    CONSTRAINT FK_Asistencia_Materia FOREIGN KEY (IdMateria) REFERENCES Materias(IdMateria)
);

-- 2. Tabla para Módulo de Exámenes
CREATE TABLE Examenes (
    IdExamen INT PRIMARY KEY IDENTITY(1,1),
    IdMateria INT,
    Parcial INT, -- 1, 2 o Final
    FechaExamen DATETIME,
    Aula NVARCHAR(20),
    CONSTRAINT FK_Examen_Materia FOREIGN KEY (IdMateria) REFERENCES Materias(IdMateria)
);

-- 3. Tabla de Configuración de Pagos (Para el Cálculo Automático)
CREATE TABLE ConfiguracionPagos (
    IdConfig INT PRIMARY KEY IDENTITY(1,1),
    TipoDocente NVARCHAR(50), -- 'Horario', 'Planta'
    PagoPorHora DECIMAL(10,2)
);

INSERT INTO ConfiguracionPagos (TipoDocente, PagoPorHora) VALUES ('Horario', 150.00);

USE JosuOsDB;
GO
ALTER TABLE Materias ADD Turno NVARCHAR(20);

USE JosuOsDB;
GO

-- Agregamos las columnas necesarias para la Fase 2
ALTER TABLE Materias ADD DiaSemana NVARCHAR(20);
ALTER TABLE Materias ADD Turno NVARCHAR(20); -- Para diferenciar Regular y Sabatino



USE JosuOsDB;
GO

-- Solo agrega 'DiaSemana' si no existe
IF NOT EXISTS (SELECT * FROM sys.columns WHERE Name = N'DiaSemana' AND Object_ID = Object_ID(N'Materias'))
BEGIN
    ALTER TABLE Materias ADD DiaSemana NVARCHAR(20);
END
GO

-- Solo agrega 'Turno' si no existe
IF NOT EXISTS (SELECT * FROM sys.columns WHERE Name = N'Turno' AND Object_ID = Object_ID(N'Materias'))
BEGIN
    ALTER TABLE Materias ADD Turno NVARCHAR(20);
END
GO

EXEC sp_help 'Materias';



USE JosuOsDB;
GO

-- Si la tabla existe pero está mal, la borramos para crearla bien
IF EXISTS (SELECT * FROM sys.objects WHERE object_id = OBJECT_ID(N'[dbo].[Examenes]') AND type in (N'U'))
BEGIN
    DROP TABLE Examenes;
END
GO

-- Creamos la tabla con los nombres EXACTOS que busca tu código
CREATE TABLE Examenes (
    IdExamen INT PRIMARY KEY IDENTITY(1,1),
    NombreDocente NVARCHAR(100),
    NombreClase NVARCHAR(100),
    NumRecibo NVARCHAR(50),
    FechaRecibo DATE,
    Estudiante NVARCHAR(100),
    TipoExamen NVARCHAR(50)
);
GO

USE JosuOsDB;
GO
ALTER TABLE Docentes ADD NivelAcademico NVARCHAR(50);
ALTER TABLE Docentes ADD Facultad NVARCHAR(100);
GO

-- Actualizamos a la profe Irene con sus datos reales para la demo
USE JosuOsDB;
GO

-- 1. Agregamos las columnas que faltan
ALTER TABLE Docentes ADD NivelAcademico NVARCHAR(50);
ALTER TABLE Docentes ADD Facultad NVARCHAR(100);
GO

-- 2. Ahora sí, le ponemos los datos a la Profe Irene para que se vea Pro en la demo
UPDATE Docentes 
SET NivelAcademico = 'Maestría', Facultad = 'Facultad de Ingeniería' 
WHERE NombreFull LIKE '%Irene Martínez%';
GO

USE JosuOsDB;
GO
-- Esto añade la columna a la tabla de materias para que el sistema pueda leerla
IF NOT EXISTS (SELECT * FROM sys.columns WHERE Name = N'Cuatrimestre' AND Object_ID = Object_ID(N'Materias'))
BEGIN
    ALTER TABLE Materias ADD Cuatrimestre NVARCHAR(50);
END
GO

ALTER TABLE Materias ADD Facultad NVARCHAR(100);
ALTER TABLE Examenes ADD Facultad NVARCHAR(100);

SELECT IdMateria, Estado, Fecha FROM Asistencias WHERE Estado = 'Presente'
