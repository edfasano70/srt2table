# Srt2Table (web)

Convierte archivos de subtítulos **.srt** a **CSV** (Excel) o **DOCX** (Word).
Aplicación web de un solo endpoint, sin base de datos y sin build step.

## Stack

- **Frontend:** HTML5 + CSS propio + **JavaScript vanilla** (drag & drop).
- **Backend:** **PHP 8** (un único `index.php`).
- **Dependencia:** `phpoffice/phpword` vía **Composer** (solo para DOCX).
- **Servidor:** Apache + HTTPS.

## Estructura

```
index.php            Página y endpoint de conversión
lib/srt.php          Parser SRT + generación de CSV/DOCX (testeable)
VERSION              Versión de la app (fuente única)
tests/run.php        Tests CLI de la conversión
tests/fixtures/      SRT de ejemplo
composer.json/lock   Dependencias
vendor/              Dependencias instaladas (no versionado)
```

## Requisitos

- PHP 8.1+ con `mbstring`.
- Composer (solo para instalar PHPWord).
- Apache con HTTPS (opcional para `php -S` local).

## Instalación

```bash
composer install          # crea vendor/ con PHPWord
```

## Desarrollo local

```bash
php -S localhost:8000     # luego abre http://localhost:8000/
```

## Tests

```bash
php tests/run.php         # 0 = OK; salida PASS/FAIL
```

## Despliegue

```bash
# 1. Subir el proyecto al servidor (vendor/ se excluye y se regenera)
rsync -az --delete --exclude '.git' --exclude 'vendor' ./ usuario@SERVIDOR:/var/www/html/srt2table/

# 2. Instalar dependencias en el servidor
ssh usuario@SERVIDOR 'cd /var/www/html/srt2table && composer install --no-dev'
```

Apache: exponer por HTTPS y **bloquear código/proyecto**. Bloque de ejemplo para
un vhost existente (los cambios requieren `sudo systemctl reload apache2`):

```apache
Alias /srt2table /var/www/html/srt2table

<Directory /var/www/html/srt2table>
    Options -Indexes +FollowSymLinks
    Require all granted
    <FilesMatch "\.(py|ini|lock|spec|json|bak)$">
        Require all denied
    </FilesMatch>
</Directory>
<Directory /var/www/html/srt2table/lib>
    Require all denied
</Directory>
<Directory /var/www/html/srt2table/tests>
    Require all denied
</Directory>
<Files "VERSION">
    Require all denied
</Files>
```

No hay base de datos: la conversión es stateless.

## Notas

- La versión mostrada bajo el título se lee de `VERSION`.
- El CSV se genera con **BOM UTF-8** para que Excel respete acentos y la flecha `→`.
