-- =====================================================================
--  PoliMusic - Usuarios de trabajo
--  Crea (si no existen) las cuentas indicadas en los requerimientos:
--      userAdmin  -> administrador (USER_TYPE = 1)
--      epn.fis    -> usuario normal (USER_TYPE = 0)
--  La contrasena se guarda como SHA-1 hexadecimal (40 caracteres), que es
--  el formato que la aplicacion genera al registrar/editar usuarios.
--  SHA1('12345678') = 7C222FB2927D828AF22F592134E8932480637C0D
--
--  Ejecucion:
--      sqlcmd -S localhost -U usr_polimusic -P Politecnica1 -C -d BDD_PoliMusic -i seed_usuarios.sql
-- =====================================================================
USE [BDD_PoliMusic];
GO

SET NOCOUNT ON;
GO

-- --------------------------- Administrador ---------------------------
IF NOT EXISTS (SELECT 1 FROM TBL_USER WHERE LOWER(USERNAME) = 'useradmin')
BEGIN
    INSERT INTO TBL_USER (USERNAME, PASSWORD, EMAIL, BIRTHDAY, USER_PHOTO, USER_TYPE)
    VALUES ('userAdmin', '7C222FB2927D828AF22F592134E8932480637C0D',
            'useradmin@epn.edu.ec', '2001-12-16', '~/UserImages/root.png', 1);
    PRINT 'Usuario userAdmin creado (ADMIN).';
END
ELSE
BEGIN
    -- Garantiza que userAdmin tenga rol de administrador.
    UPDATE TBL_USER SET USER_TYPE = 1 WHERE LOWER(USERNAME) = 'useradmin' AND USER_TYPE <> 1;
    PRINT 'El usuario userAdmin ya existe.';
END
GO

-- ------------------------- Usuario normal ---------------------------
IF NOT EXISTS (SELECT 1 FROM TBL_USER WHERE LOWER(USERNAME) = 'epn.fis')
BEGIN
    INSERT INTO TBL_USER (USERNAME, PASSWORD, EMAIL, BIRTHDAY, USER_PHOTO, USER_TYPE)
    VALUES ('epn.fis', '7C222FB2927D828AF22F592134E8932480637C0D',
            'epn.fis@epn.edu.ec', '2000-01-01', '~/UserImages/default.jpg', 0);
    PRINT 'Usuario epn.fis creado (USUARIO NORMAL).';
END
ELSE
BEGIN
    PRINT 'El usuario epn.fis ya existe.';
END
GO

-- ------------------------- Usuario normal 2 -------------------------
IF NOT EXISTS (SELECT 1 FROM TBL_USER WHERE LOWER(USERNAME) = 'usernormal')
BEGIN
    INSERT INTO TBL_USER (USERNAME, PASSWORD, EMAIL, BIRTHDAY, USER_PHOTO, USER_TYPE)
    VALUES ('userNormal', '7C222FB2927D828AF22F592134E8932480637C0D',
            'usrnormal@epn.edu.ec', '2000-12-15', '~/UserImages/default.jpg', 0);
    PRINT 'Usuario userNormal creado (USUARIO NORMAL).';
END
ELSE
BEGIN
    PRINT 'El usuario userNormal ya existe.';
END
GO

SELECT ID_USER, USERNAME, EMAIL, USER_TYPE,
       CASE WHEN LEN(PASSWORD) = 40 THEN 'SHA-1' ELSE 'TEXTO PLANO' END AS FORMATO
FROM TBL_USER
ORDER BY ID_USER;
GO
