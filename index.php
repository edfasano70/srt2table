<?php

error_reporting(E_ALL);
ini_set('display_errors', '0');
ini_set('display_startup_errors', '0');

// --- Versión (fuente única: fichero VERSION) ---
$appVersion = trim((string) @file_get_contents(__DIR__ . '/VERSION'));
if ($appVersion === '') {
    $appVersion = '0.0.0';
}

// --- Límite de subida ---
// Por defecto se usa el menor entre upload_max_filesize y post_max_size de PHP.
// Pon un valor > 0 aquí para imponer un tope propio en MB.
$appMaxUploadMb = 0;

function iniSizeToBytes($value) {
    $value = trim((string) $value);
    if ($value === '' || $value === '-1') {
        return 0;
    }
    $unit = strtolower(substr($value, -1));
    $num = (float) $value;
    switch ($unit) {
        case 'g': $num *= 1024 * 1024 * 1024; break;
        case 'm': $num *= 1024 * 1024; break;
        case 'k': $num *= 1024; break;
    }
    return (int) $num;
}

function effectiveUploadLimitBytes($overrideMb = 0) {
    $candidates = [];
    foreach (['upload_max_filesize', 'post_max_size'] as $key) {
        $bytes = iniSizeToBytes(ini_get($key));
        if ($bytes > 0) {
            $candidates[] = $bytes;
        }
    }
    if ($overrideMb > 0) {
        $candidates[] = (int) $overrideMb * 1024 * 1024;
    }
    return $candidates ? min($candidates) : 0;
}

function humanSize($bytes) {
    if ($bytes <= 0) {
        return 'sin límite';
    }
    $units = ['B', 'KB', 'MB', 'GB'];
    $size = (float) $bytes;
    $i = 0;
    while ($size >= 1024 && $i < count($units) - 1) {
        $size /= 1024;
        $i++;
    }
    return round($size, $i === 0 ? 0 : 1) . ' ' . $units[$i];
}

$maxUploadBytes = effectiveUploadLimitBytes($appMaxUploadMb);

// --- Dependencias ---
if (file_exists(__DIR__ . '/vendor/autoload.php')) {
    require_once __DIR__ . '/vendor/autoload.php';
}
require_once __DIR__ . '/lib/srt.php';

$phpWordInstalled = class_exists('\PhpOffice\PhpWord\PhpWord');

// --- Cabeceras de seguridad base ---
header('X-Content-Type-Options: nosniff');
header('Referrer-Policy: no-referrer');

// --- Lógica principal ---
$error = '';
$isPost = (($_SERVER['REQUEST_METHOD'] ?? '') === 'POST');

if ($isPost) {
    $file = $_FILES['srt_file'] ?? null;
    $output_format = (string) ($_POST['output_format'] ?? 'csv');
    $content_length = (int) ($_SERVER['CONTENT_LENGTH'] ?? 0);

    if (empty($_POST) && empty($_FILES) && $content_length > 0) {
        // El navegador cortó la subida: casi siempre por superar post_max_size.
        $error = 'No se pudo procesar la subida (demasiado grande para el límite del servidor).';
    } elseif (!is_array($file) || ($file['error'] ?? UPLOAD_ERR_NO_FILE) === UPLOAD_ERR_NO_FILE) {
        $error = 'Elige un archivo .srt.';
    } elseif (in_array($file['error'], [UPLOAD_ERR_INI_SIZE, UPLOAD_ERR_FORM_SIZE], true)
        || ($maxUploadBytes > 0 && (int) $file['size'] > $maxUploadBytes)) {
        $error = 'El archivo es demasiado grande. Máximo permitido: ' . humanSize($maxUploadBytes) . '.';
    } elseif ($file['error'] !== UPLOAD_ERR_OK) {
        $error = 'Falló la subida (código ' . (int) $file['error'] . '). Inténtalo de nuevo.';
    } elseif (strtolower(pathinfo((string) $file['name'], PATHINFO_EXTENSION)) !== 'srt') {
        $error = 'Solo se aceptan archivos .srt.';
    } elseif (!in_array($output_format, ['csv', 'docx'], true)) {
        $error = 'Formato de salida no válido.';
    } elseif ($output_format === 'docx' && !$phpWordInstalled) {
        $error = 'La exportación a DOCX no está disponible: la librería PHPWord no está instalada.';
    } else {
        $raw_content = file_get_contents($file['tmp_name']);
        $encoding = mb_detect_encoding($raw_content, ['UTF-8', 'ISO-8859-1', 'Windows-1252'], true) ?: 'UTF-8';
        $srt_content = mb_convert_encoding($raw_content, 'UTF-8', $encoding);

        $project_name = trim((string) ($_POST['project_name'] ?? ''));
        if ($project_name === '') {
            $project_name = pathinfo((string) $file['name'], PATHINFO_FILENAME);
        }

        $base = pathinfo((string) $file['name'], PATHINFO_FILENAME);
        $base = preg_replace('/[^\p{L}\p{N} _.\-]+/u', '_', $base);
        $base = trim($base, " ._-");
        if ($base === '') {
            $base = 'subtitles';
        }
        $download_name = $base . '.' . $output_format;

        try {
            if ($output_format === 'csv') {
                $body = srtToCsv($srt_content);
                $mime = 'text/csv; charset=utf-8';
            } else {
                $body = srtToDocx($srt_content, $project_name);
                $mime = 'application/vnd.openxmlformats-officedocument.wordprocessingml.document';
            }

            header('Content-Type: ' . $mime);
            header('Content-Disposition: attachment; filename="' . $download_name . '"');
            header('Content-Length: ' . strlen($body));
            header('Cache-Control: no-store');
            echo $body;
            exit;
        } catch (\Throwable $e) {
            $error = 'La conversión falló: ' . $e->getMessage();
        }
    }
}

// --- Página HTML ---
$cspNonce = base64_encode(random_bytes(16));
header("Content-Security-Policy: default-src 'none'; img-src 'self'; manifest-src 'self'; style-src 'unsafe-inline'; script-src 'nonce-$cspNonce'; form-action 'self'; base-uri 'none'; frame-ancestors 'none'");
header('X-Frame-Options: DENY');
?>
<!DOCTYPE html>
<html lang="es">
<head>
    <meta charset="UTF-8">
    <meta name="viewport" content="width=device-width, initial-scale=1.0">
    <title>Srt2Table - Convierte SRT a CSV o DOCX</title>

    <link rel="icon" href="favicon.ico" sizes="any">
    <link rel="icon" type="image/png" sizes="32x32" href="favicon-32x32.png">
    <link rel="icon" type="image/png" sizes="16x16" href="favicon-16x16.png">
    <link rel="apple-touch-icon" sizes="180x180" href="apple-touch-icon.png">
    <link rel="manifest" href="site.webmanifest">

    <style>
        body { font-family: -apple-system, BlinkMacSystemFont, "Segoe UI", Roboto, "Helvetica Neue", Arial, sans-serif; line-height: 1.6; color: #333; max-width: 650px; margin: 40px auto; padding: 0 20px; background-color: #f8f9fa; }
        .container { background-color: #fff; padding: 30px; border-color: #000; border-style: solid; border-width: 1px; border-radius: 8px; box-shadow: 0 4px 6px rgba(0, 0, 0, 0.42); }
        h1 { color: #212529; text-align: center; }
        .form-group { margin-bottom: 20px; }
        label { display: block; font-weight: bold; margin-bottom: 8px; }
        input[type="text"], input[type="file"] { width: 100%; padding: 10px; border: 1px solid #ced4da; border-radius: 4px; box-sizing: border-box; }
        .radio-group label { display: inline-block; margin-right: 20px; font-weight: normal; }
        .btn { display: block; width: 100%; background-color: #007bff; color: white; padding: 12px; border: none; border-radius: 4px; font-size: 16px; cursor: pointer; text-align: center; }
        .btn:hover { background-color: #0056b3; }
        .alert { padding: 15px; margin-bottom: 20px; border: 1px solid transparent; border-radius: 4px; }
        .alert-warning { color: #856404; background-color: #fff3cd; border-color: #ffeeba; }
        .alert-danger { color: #721c24; background-color: #f8d7da; border-color: #f5c6cb; }
        .version { text-align: center; color: #6c757d; font-size: 0.8em; margin: -10px 0 20px; }
        .hint { color: #6c757d; font-size: 0.85em; margin-top: 6px; }
        .drop-zone { border: 2px dashed #adb5bd; border-radius: 6px; padding: 24px 16px; text-align: center; color: #6c757d; cursor: pointer; transition: border-color .15s, background-color .15s, color .15s; }
        .drop-zone:hover, .drop-zone:focus { border-color: #007bff; color: #0056b3; outline: none; }
        .drop-zone.dragover { border-color: #007bff; background-color: #e7f1ff; color: #0056b3; }
        .drop-zone.has-error { border-color: #dc3545; color: #dc3545; }
        .drop-zone .file-name { display: block; margin-top: 10px; color: #212529; font-weight: bold; word-break: break-all; }
        .app-logo { display: block; width: 128px; height: 128px; margin: 0 auto 10px; }
        .sr-only { position: absolute; width: 1px; height: 1px; padding: 0; margin: -1px; overflow: hidden; clip: rect(0, 0, 0, 0); white-space: nowrap; border: 0; }
        footer { text-align: center; margin-top: 30px; font-size: 0.9em; color: #6c757d; }
    </style>
</head>
<body>

<div class="container">
    <img class="app-logo" src="splash2.png" alt="Srt2Table">
    <h1>Conversor Srt2Table</h1>
    <p class="version">v<?php echo htmlspecialchars($appVersion); ?></p>

    <?php if ($error !== ''): ?>
        <div class="alert alert-danger"><?php echo htmlspecialchars($error); ?></div>
    <?php endif; ?>

    <?php if (!$phpWordInstalled): ?>
        <div class="alert alert-warning">
            <strong>Aviso:</strong> no se encontró la librería 'PHPWord'. La exportación a DOCX no funcionará. Ejecuta <code>composer install</code> en el directorio del proyecto.
        </div>
    <?php endif; ?>

    <form action="index.php" method="post" enctype="multipart/form-data" id="convert-form">
        <div class="form-group">
            <label for="srt_file">1. Selecciona el archivo SRT</label>
            <div id="drop-zone" class="drop-zone" role="button" tabindex="0" aria-controls="srt_file">
                Arrastra y suelta tu archivo <strong>.srt</strong> aquí<br>o haz clic para seleccionarlo
                <span id="file-name" class="file-name"></span>
            </div>
            <input type="file" id="srt_file" name="srt_file" accept=".srt" required class="sr-only">
            <p class="hint">Tamaño máximo: <?php echo htmlspecialchars(humanSize($maxUploadBytes)); ?></p>
        </div>

        <div class="form-group">
            <label for="project_name">2. Nombre del proyecto (título del DOCX)</label>
            <input type="text" id="project_name" name="project_name" placeholder="Se deduce del nombre del archivo si se deja en blanco">
        </div>

        <div class="form-group">
            <label>3. Elige el formato de salida</label>
            <div class="radio-group">
                <label>
                    <input type="radio" name="output_format" value="csv" checked> CSV (para Excel)
                </label>
                <label>
                    <input type="radio" name="output_format" value="docx" <?php if (!$phpWordInstalled) echo 'disabled'; ?>> DOCX (para Word)
                </label>
            </div>
        </div>

        <button type="submit" class="btn">Convertir y descargar</button>
    </form>
</div>

<footer>
    Idea de Pilar Terán &bull; Programado por @EdFasano70
</footer>

<script nonce="<?php echo $cspNonce; ?>">
    (function () {
        const dropZone = document.getElementById('drop-zone');
        const fileInput = document.getElementById('srt_file');
        const fileName = document.getElementById('file-name');
        const projectNameInput = document.getElementById('project_name');

        function autoFillProjectName(filename) {
            const basename = filename.substring(0, filename.lastIndexOf('.')) || filename;
            projectNameInput.value = basename;
        }

        function setFile(file) {
            if (!file) return;
            if (!/\.srt$/i.test(file.name)) {
                dropZone.classList.add('has-error');
                fileName.textContent = 'Solo se permiten archivos .srt: ' + file.name;
                fileInput.value = '';
                return;
            }
            dropZone.classList.remove('has-error');
            const dt = new DataTransfer();
            dt.items.add(file);
            fileInput.files = dt.files;
            fileName.textContent = file.name;
            autoFillProjectName(file.name);
        }

        fileInput.addEventListener('change', function () {
            if (this.files.length > 0) {
                setFile(this.files[0]);
            }
        });

        dropZone.addEventListener('click', function () { fileInput.click(); });
        dropZone.addEventListener('keydown', function (e) {
            if (e.key === 'Enter' || e.key === ' ') {
                e.preventDefault();
                fileInput.click();
            }
        });

        ['dragenter', 'dragover'].forEach(function (ev) {
            dropZone.addEventListener(ev, function (e) {
                e.preventDefault();
                e.stopPropagation();
                dropZone.classList.add('dragover');
            });
        });
        ['dragleave', 'dragend'].forEach(function (ev) {
            dropZone.addEventListener(ev, function (e) {
                e.preventDefault();
                e.stopPropagation();
                dropZone.classList.remove('dragover');
            });
        });
        dropZone.addEventListener('drop', function (e) {
            e.preventDefault();
            e.stopPropagation();
            dropZone.classList.remove('dragover');
            if (e.dataTransfer.files && e.dataTransfer.files.length > 0) {
                setFile(e.dataTransfer.files[0]);
            }
        });

        // Evita que el navegador abra el archivo si se suelta fuera de la zona
        ['dragover', 'drop'].forEach(function (ev) {
            window.addEventListener(ev, function (e) { e.preventDefault(); });
        });
    })();
</script>

</body>
</html>
