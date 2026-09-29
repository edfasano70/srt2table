<?php
/**
 * Srt2Table - Conversión de SRT a CSV/DOCX.
 *
 * Funciones puras (salvo srtToDocx, que requiere PHPWord). Sin efectos de
 * salida: devuelven strings para poder testearlas con tests/run.php.
 */

/**
 * Elimina caracteres de control inválidos en XML, conservando tab, LF y CR.
 */
function removeInvalidXmlChars($text) {
    return preg_replace('/[^\x{0009}\x{000a}\x{000d}\x{0020}-\x{D7FF}\x{E000}-\x{FFFD}\x{10000}-\x{10FFFF}]/u', '', (string) $text);
}

/**
 * Cierra las etiquetas HTML soportadas (i, b, u, font) que queden abiertas.
 */
function fixHtmlTags($text) {
    preg_match_all('/<(\/?)(\w+)[^>]*>/', $text, $matches, PREG_SET_ORDER);

    $stack = [];
    $supportedTags = ['i', 'b', 'u', 'font'];

    foreach ($matches as $match) {
        $isClosingTag = $match[1] === '/';
        $tagName = strtolower($match[2]);

        if (!in_array($tagName, $supportedTags, true)) {
            continue;
        }

        if ($isClosingTag) {
            if (!empty($stack) && end($stack) === $tagName) {
                array_pop($stack);
            }
        } else {
            $stack[] = $tagName;
        }
    }

    while (!empty($stack)) {
        $tagToClose = array_pop($stack);
        $text .= "</{$tagToClose}>";
    }

    return $text;
}

/**
 * Añade texto con etiquetas simples (<i>, <b>) a un TextRun de PHPWord.
 *
 * @param \PhpOffice\PhpWord\Element\TextRun $element
 * @param string $textWithTags
 */
function addFormattedText($element, $textWithTags) {
    $fragments = preg_split('/(<[^>]+>)/', $textWithTags, -1, PREG_SPLIT_DELIM_CAPTURE);

    $styleStack = [];
    $defaultStyle = ['italic' => false, 'bold' => false];

    foreach ($fragments as $fragment) {
        if ($fragment === '' || $fragment === null) {
            continue;
        }

        if (preg_match('/<(\/?)(\w+)[^>]*>/', $fragment, $matches)) {
            $isClosingTag = ($matches[1] === '/');
            $tagName = strtolower($matches[2]);

            if ($isClosingTag) {
                if (!empty($styleStack) && end($styleStack) === $tagName) {
                    array_pop($styleStack);
                }
            } else {
                if (in_array($tagName, ['i', 'b'], true)) {
                    $styleStack[] = $tagName;
                }
            }
        } else {
            $currentStyle = $defaultStyle;
            $currentStyle['italic'] = in_array('i', $styleStack, true);
            $currentStyle['bold'] = in_array('b', $styleStack, true);

            $element->addText(htmlspecialchars($fragment), $currentStyle);
        }
    }
}

/**
 * Extrae subtítulos de un contenido SRT.
 *
 * Parser tolerante: acepta BOM, saltos CRLF/CR, bloques sin línea en blanco,
 * milisegundos con '.' o ',', milisegundos opcionales y etiquetas sin cerrar.
 *
 * @return array<array{start:string,end:string,dialogue:string}>
 */
function parseSrt($srtContent) {
    // Patrón de timecode: HH:MM:SS[.,mmm] --> HH:MM:SS[.,mmm]
    $timecodeRe = '/(\d{1,2}:\d{2}:\d{2})(?:[,.]\d{1,3})?\s*-->\s*(\d{1,2}:\d{2}:\d{2})(?:[,.]\d{1,3})?/';

    $content = str_replace(["\r\n", "\r"], "\n", (string) $srtContent);
    $content = removeInvalidXmlChars($content);
    $content = preg_replace('/^\x{FEFF}|\x{FEFF}/u', '', $content); // BOM UTF-8
    $lines = explode("\n", $content);
    $count = count($lines);

    $subtitles = [];
    $current = null;

    for ($i = 0; $i < $count; $i++) {
        $line = rtrim($lines[$i], " \t");

        if (preg_match($timecodeRe, $line, $m)) {
            if ($current !== null) {
                $subtitles[] = $current;
            }
            $current = ['start' => $m[1], 'end' => $m[2], 'dialogue' => ''];
            continue;
        }

        if ($current === null) {
            continue;
        }

        // Una línea que es solo el índice del siguiente bloque no es diálogo.
        if (preg_match('/^\d+$/', trim($line))) {
            $j = $i + 1;
            while ($j < $count && trim($lines[$j]) === '') {
                $j++;
            }
            if ($j < $count && preg_match($timecodeRe, trim($lines[$j]))) {
                $subtitles[] = $current;
                $current = null;
                continue;
            }
        }

        $text = trim($line);
        if ($text === '') {
            continue;
        }
        $current['dialogue'] .= ($current['dialogue'] === '' ? '' : ' ') . $text;
    }

    if ($current !== null) {
        $subtitles[] = $current;
    }

    foreach ($subtitles as &$sub) {
        $sub['dialogue'] = fixHtmlTags(trim($sub['dialogue']));
    }
    unset($sub);

    return $subtitles;
}

/**
 * Convierte SRT a CSV (UTF-8 con BOM para compatibilidad con Excel).
 *
 * @return string
 */
function srtToCsv($srtContent) {
    $subtitles = parseSrt($srtContent);

    $output = fopen('php://temp', 'r+');
    fwrite($output, "\xEF\xBB\xBF"); // BOM UTF-8
    fputcsv($output, ['Inicio', 'Flecha', 'Fin', 'Personaje', 'Diálogo']);

    foreach ($subtitles as $sub) {
        fputcsv($output, [
            $sub['start'],
            '→',
            $sub['end'],
            '',
            strip_tags($sub['dialogue']),
        ]);
    }

    rewind($output);
    $csv = stream_get_contents($output);
    fclose($output);

    return $csv;
}

/**
 * Convierte SRT a un documento DOCX y devuelve su contenido binario.
 *
 * @return string
 * @throws RuntimeException si PHPWord no está instalado.
 */
function srtToDocx($srtContent, $projectName) {
    if (!class_exists('\PhpOffice\PhpWord\PhpWord')) {
        throw new RuntimeException("PHPWord no está instalado. Ejecuta 'composer install'.");
    }

    $subtitles = parseSrt($srtContent);
    $phpWord = new \PhpOffice\PhpWord\PhpWord();

    $phpWord->setDefaultFontName('Arial');
    $section = $phpWord->addSection([
        'marginLeft' => \PhpOffice\PhpWord\Shared\Converter::cmToTwip(1),
        'marginRight' => \PhpOffice\PhpWord\Shared\Converter::cmToTwip(1),
    ]);

    $footer = $section->addFooter();
    $footer->addText("Generado por Srt2Table", null, ['alignment' => \PhpOffice\PhpWord\SimpleType\Jc::CENTER]);

    $titleFont = ['size' => 20, 'color' => '87CEEB', 'bold' => true];
    $titleParagraph = ['alignment' => \PhpOffice\PhpWord\SimpleType\Jc::CENTER];
    $section->addText($projectName ?: 'Subtitle Table', $titleFont, $titleParagraph);
    $section->addTextBreak(1);

    $table = $section->addTable([
        'width' => 100 * 50,
        'unit' => \PhpOffice\PhpWord\SimpleType\TblWidth::PERCENT,
        'borderSize' => 0,
        'cellMargin' => 80,
    ]);

    $headerStyle = ['bold' => true];
    $paragraphCenter = ['alignment' => \PhpOffice\PhpWord\SimpleType\Jc::CENTER, 'spaceAfter' => 0, 'spaceBefore' => 0];

    $table->addRow();
    $table->addCell(2500)->addText('Código de tiempo', $headerStyle, $paragraphCenter);
    $table->addCell(1500)->addText('Personaje', $headerStyle, $paragraphCenter);
    $table->addCell(7000)->addText('Diálogo', $headerStyle, $paragraphCenter);

    $paragraphLeft = ['spaceAfter' => 0, 'spaceBefore' => 0];

    foreach ($subtitles as $sub) {
        $table->addRow();
        $timecode = $sub['start'] . ' → ' . $sub['end'];

        $table->addCell(3000)->addText($timecode, null, $paragraphCenter);
        $table->addCell(2000)->addText('', null, $paragraphCenter);

        $dialogueCell = $table->addCell(5000);
        $textRun = $dialogueCell->addTextRun($paragraphLeft);
        addFormattedText($textRun, $sub['dialogue']);
    }

    $objWriter = \PhpOffice\PhpWord\IOFactory::createWriter($phpWord, 'Word2007');

    ob_start();
    $objWriter->save('php://output');

    return ob_get_clean();
}
