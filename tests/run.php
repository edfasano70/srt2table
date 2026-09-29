<?php
/**
 * Tests de la conversión SRT. Uso: php tests/run.php
 */

$autoload = __DIR__ . '/../vendor/autoload.php';
if (file_exists($autoload)) {
    require $autoload;
}
require __DIR__ . '/../lib/srt.php';

$failures = 0;

function check(string $name, $cond): void {
    global $failures;
    if ($cond) {
        echo "PASS  $name\n";
    } else {
        echo "FAIL  $name\n";
        $failures++;
    }
}

function fixture(string $name): string {
    return file_get_contents(__DIR__ . '/fixtures/' . $name);
}

$basic = parseSrt(fixture('basic.srt'));
check('basic: 2 subtítulos', count($basic) === 2);
check('basic: timecode inicio', ($basic[0]['start'] ?? '') === '00:00:01');
check('basic: timecode fin', ($basic[0]['end'] ?? '') === '00:00:03');
check('basic: diálogo simple', ($basic[0]['dialogue'] ?? '') === 'Hello world');
check('basic: une líneas y conserva etiquetas', ($basic[1]['dialogue'] ?? '') === '<i>Second</i> line continued here');

$tolerant = parseSrt("\xEF\xBB\xBF" . fixture('tolerant.srt'));
check('tolerant: BOM + sin líneas en blanco => 2', count($tolerant) === 2);
check('tolerant: ms con punto', ($tolerant[0]['start'] ?? '') === '00:00:01');
check('tolerant: cierra etiqueta abierta', ($tolerant[0]['dialogue'] ?? '') === '<i>Uno</i>');
check('tolerant: segundo bloque', ($tolerant[1]['dialogue'] ?? '') === 'Dos');

$csv = srtToCsv(fixture('basic.srt'));
check('csv: BOM UTF-8', substr($csv, 0, 3) === "\xEF\xBB\xBF");
$csvRows = explode("\n", trim(substr($csv, 3)));
check('csv: cabecera', ($csvRows[0] ?? '') === 'Inicio,Flecha,Fin,Personaje,Diálogo');
check('csv: 2 filas de datos', count($csvRows) === 3);
check('csv: sin etiquetas y con texto', strpos($csvRows[2], '<i>') === false && strpos($csvRows[2], 'Second line continued here') !== false);

if (class_exists('\PhpOffice\PhpWord\PhpWord')) {
    $docx = srtToDocx(fixture('basic.srt'), 'Prueba');
    check('docx: firma ZIP', substr($docx, 0, 2) === 'PK');
} else {
    echo "SKIP  docx (PHPWord no instalado)\n";
}

echo "\n" . ($failures === 0 ? 'OK - todos los tests pasan' : $failures . ' test(s) fallaron') . "\n";
exit($failures === 0 ? 0 : 1);
