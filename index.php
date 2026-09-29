<?php

/* Pasos para Poner en Marcha la Versión Web:
Requisitos: 
    Necesitarás un servidor web con PHP (versión 7.4 o superior es recomendable) y Composer instalado.
    Crear Archivos: Guarda los dos bloques de código que te proporciono a continuación como composer.json e index.php en un directorio de tu servidor web.
    Instalar Dependencias: 
        Abre una terminal en ese directorio y ejecuta el comando composer install. 
        Esto descargará automáticamente la librería phpoffice/phpword y creará una carpeta vendor.
Una vez completado, podrás acceder a index.php desde tu navegador y la aplicación estará lista para usar. */

error_reporting(E_ALL);
ini_set('display_errors', '0');
ini_set('display_startup_errors', '0');

// --- Dependency Check and Autoloader ---
// This handles loading the PHPWord library installed via Composer.
if (file_exists(__DIR__ . '/vendor/autoload.php')) {
    require_once __DIR__ . '/vendor/autoload.php';
}

/**
 * Parses text with simple HTML tags (<i>, <b>) and adds it to a PHPWord TextRun
 * with corresponding formatting.
 *
 * @param \PhpOffice\PhpWord\Element\TextRun|\PhpOffice\PhpWord\Element\Cell $element The parent element (TextRun or Cell).
 * @param string $textWithTags The text containing HTML-like tags.
 */
function addFormattedText($element, $textWithTags) {
    // This function is designed to work with a TextRun.
    // The previous check for a Cell was problematic and has been removed for clarity and safety.
    // The calling code should ensure it passes a TextRun.

    $fragments = preg_split('/(<[^>]+>)/', $textWithTags, -1, PREG_SPLIT_DELIM_CAPTURE);

    // Use a stack to correctly handle nested tags
    $styleStack = [];
    $defaultStyle = ['italic' => false, 'bold' => false];

    foreach ($fragments as $fragment) {
        if (empty($fragment)) {
            continue;
        }

        // Check if the fragment is a tag
        if (preg_match('/<(\/?)(\w+)[^>]*>/', $fragment, $matches)) {
            $isClosingTag = ($matches[1] === '/');
            $tagName = strtolower($matches[2]);

            if ($isClosingTag) {
                // On a closing tag, pop the last style if it matches
                if (!empty($styleStack) && end($styleStack) === $tagName) {
                    array_pop($styleStack);
                }
            } else {
                // On an opening tag, push it to the stack
                if (in_array($tagName, ['i', 'b'])) {
                    $styleStack[] = $tagName;
                }
            }
        } else {
            // It's a text fragment. Calculate the current style from the stack.
            $currentStyle = $defaultStyle;
            $currentStyle['italic'] = in_array('i', $styleStack);
            $currentStyle['bold'] = in_array('b', $styleStack);

            // Add the text with the calculated style
            $element->addText(htmlspecialchars($fragment), $currentStyle);
        }
    }
}

/**
 * Analyzes text and closes any open HTML tags (i, b, u, font).
 * Uses a stack to correctly handle nested tags.
 *
 * @param string $text The input text with potentially unclosed tags.
 * @return string The corrected text with all tags properly closed.
 */
function fixHtmlTags($text) {
    preg_match_all('/<(\/?)(\w+)[^>]*>/', $text, $matches, PREG_SET_ORDER);

    $stack = [];
    $supportedTags = ['i', 'b', 'u', 'font'];

    foreach ($matches as $match) {
        $isClosingTag = $match[1] === '/';
        $tagName = strtolower($match[2]);

        if (!in_array($tagName, $supportedTags)) {
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
 * Removes invalid XML control characters, preserving newlines and tabs.
 *
 * @param string $text The text to clean.
 * @return string The cleaned text.
 */
function removeInvalidXmlChars($text) {
    // This regex removes characters that are invalid in XML, except for tab, newline, and carriage return.
    return preg_replace('/[^\x{0009}\x{000a}\x{000d}\x{0020}-\x{D7FF}\x{E000}-\x{FFFD}\x{10000}-\x{10FFFF}]/u', '', $text);
}

/**
 * Parses the content of an SRT file into a structured array.
 *
 * @param string $srtContent The full content of the SRT file.
 * @return array An array of subtitle blocks.
 */
function parseSrt($srtContent) {
    $subtitles = [];
    // Normalize line endings and clean content
    $content = str_replace("\r\n", "\n", $srtContent);
    $content = removeInvalidXmlChars($content);
    $blocks = preg_split("/\n\n+/", trim($content));

    foreach ($blocks as $block) {
        $lines = explode("\n", trim($block));
        if (count($lines) < 2) {
            continue;
        }

        // Time line
        if (preg_match('/(\d{2}:\d{2}:\d{2}),\d+\s+-->\s+(\d{2}:\d{2}:\d{2}),\d+/', $lines[1], $matches)) {
            $startTime = $matches[1];
            $endTime = $matches[2];

            // Dialogue
            $dialogueLines = array_slice($lines, 2);
            $rawDialogue = implode(' ', array_map('trim', $dialogueLines));
            $dialogue = fixHtmlTags($rawDialogue);

            $subtitles[] = [
                'start' => $startTime,
                'end' => $endTime,
                'dialogue' => $dialogue,
            ];
        }
    }
    return $subtitles;
}

/**
 * Converts SRT content to CSV and streams it for download.
 *
 * @param string $srtContent The content of the SRT file.
 * @param string $outputFilename The desired filename for the download.
 */
function convertSrtToCsv($srtContent, $outputFilename) {
    header('Content-Type: text/csv; charset=utf-8');
    header('Content-Disposition: attachment; filename="' . $outputFilename . '"');

    $subtitles = parseSrt($srtContent);
    $output = fopen('php://output', 'w');

    // Header row
    fputcsv($output, ['Start', 'Arrow', 'End', 'Character', 'Dialogue']);

    foreach ($subtitles as $sub) {
        fputcsv($output, [
            $sub['start'],
            '→',
            $sub['end'],
            '', // Character
            strip_tags($sub['dialogue']) // CSV doesn't support formatting
        ]);
    }
    fclose($output);
}

/**
 * Converts SRT content to DOCX and streams it for download.
 *
 * @param string $srtContent The content of the SRT file.
 * @param string $outputFilename The desired filename for the download.
 * @param string $projectName The title for the document.
 */
function convertSrtToDocx($srtContent, $outputFilename, $projectName) {
    if (!class_exists('\PhpOffice\PhpWord\PhpWord')) {
        die("Error: The 'PHPWord' library is not available. Please run 'composer install'.");
    }

    $subtitles = parseSrt($srtContent);
    $phpWord = new \PhpOffice\PhpWord\PhpWord();

    // --- Page and font configuration ---
    $phpWord->setDefaultFontName('Arial');
    $section = $phpWord->addSection([
        'marginLeft' => \PhpOffice\PhpWord\Shared\Converter::cmToTwip(1),
        'marginRight' => \PhpOffice\PhpWord\Shared\Converter::cmToTwip(1),
    ]);

    // --- Footer ---
    $footer = $section->addFooter();
    $footer->addText("Generated by Srt2Table", null, ['alignment' => \PhpOffice\PhpWord\SimpleType\Jc::CENTER]);

    // --- Document Title ---
    $titleFont = ['size' => 20, 'color' => '87CEEB', 'bold' => true]; // Azul celeste, tamaño 20
    $titleParagraph = ['alignment' => \PhpOffice\PhpWord\SimpleType\Jc::CENTER];
    $section->addText($projectName ?: 'Subtitle Table', $titleFont, $titleParagraph);
    // Add a space after the title before the table starts
    $section->addTextBreak(1);

    // --- Table Creation ---
    $header = ['Timecode', 'Character', 'Dialogue'];
    $table = $section->addTable([
        'width' => 100 * 50, // 100% in fiftieths of a percent
        'unit' => \PhpOffice\PhpWord\SimpleType\TblWidth::PERCENT,
        'borderSize' => 0, // This removes all borders, internal and external.
        'cellMargin' => 80
    ]);

    // Header row
    $table->addRow();
    $headerStyle = ['bold' => true];
    $paragraphCenter = ['alignment' => \PhpOffice\PhpWord\SimpleType\Jc::CENTER, 'spaceAfter' => 0, 'spaceBefore' => 0];
    $table->addCell(2500)->addText($header[0], $headerStyle, $paragraphCenter);
    $table->addCell(1500)->addText($header[1], $headerStyle, $paragraphCenter);
    $table->addCell(7000)->addText($header[2], $headerStyle, $paragraphCenter);

    // --- Table Population ---
    $paragraphLeft = ['spaceAfter' => 0, 'spaceBefore' => 0];

    foreach ($subtitles as $sub) {
        $table->addRow();
        $timecode = $sub['start'] . ' → ' . $sub['end'];

        // Column 1: Timecode
        $table->addCell(3000)->addText($timecode, null, $paragraphCenter);
        // Column 2: Character
        $table->addCell(2000)->addText('', null, $paragraphCenter);
        // Column 3: Dialogue
        $dialogueCell = $table->addCell(5000);
        $textRun = $dialogueCell->addTextRun($paragraphLeft);
        addFormattedText($textRun, $sub['dialogue']);
    }

    // --- Stream the file for download ---
    header('Content-Type: application/vnd.openxmlformats-officedocument.wordprocessingml.document');
    header('Content-Disposition: attachment;filename="' . $outputFilename . '"');
    header('Cache-Control: max-age=0');

    $objWriter = \PhpOffice\PhpWord\IOFactory::createWriter($phpWord, 'Word2007');
    $objWriter->save('php://output');
}

// --- Main Execution Logic ---
if ($_SERVER['REQUEST_METHOD'] === 'POST' && isset($_FILES['srt_file']) && $_FILES['srt_file']['error'] === UPLOAD_ERR_OK) {
    $srt_tmp_path = $_FILES['srt_file']['tmp_name'];
    $original_filename = $_FILES['srt_file']['name'];

    $project_name = $_POST['project_name'] ?? pathinfo($original_filename, PATHINFO_FILENAME);
    $output_format = $_POST['output_format'] ?? 'csv';

    // Detect encoding and read content
    $raw_content = file_get_contents($srt_tmp_path);
    $encoding = mb_detect_encoding($raw_content, 'UTF-8, ISO-8859-1, Windows-1252', true);
    $srt_content = mb_convert_encoding($raw_content, 'UTF-8', $encoding);

    $base_filename = pathinfo($original_filename, PATHINFO_FILENAME);
    $output_filename = $base_filename . '.' . $output_format;

    if ($output_format === 'csv') {
        convertSrtToCsv($srt_content, $output_filename);
    } elseif ($output_format === 'docx') {
        convertSrtToDocx($srt_content, $output_filename, $project_name);
    }
    exit; // Stop script execution after file download
}

$phpWordInstalled = class_exists('\PhpOffice\PhpWord\PhpWord');

?>
<!DOCTYPE html>
<html lang="en">
<head>
    <meta charset="UTF-8">
    <meta name="viewport" content="width=device-width, initial-scale=1.0">
    <title>Srt2Table - Convert SRT to CSV or DOCX</title>

<!--     <link rel="apple-touch-icon" sizes="180x180" href="/apple-touch-icon.png">
    <link rel="icon" type="image/png" sizes="32x32" href="/favicon-32x32.png">
    <link rel="icon" type="image/png" sizes="16x16" href="/favicon-16x16.png">
    <link rel="manifest" href="/site.webmanifest"> -->

    <link rel="icon" href="favicon.ico" type="image/x-icon">

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
        footer { text-align: center; margin-top: 30px; font-size: 0.9em; color: #6c757d; }
    </style>
</head>
<body>

<div class="container">
    <div style="text-align: center; margin-bottom: 20px width=100%"><img src="splash.png">
    </div>
    <h1>Srt2Table Converter</h1>

    <?php if (!$phpWordInstalled): ?>
        <div class="alert alert-warning">
            <strong>Warning:</strong> The 'PHPWord' library is not found. DOCX export will not work. Please run <code>composer install</code> in the project directory.
        </div>
    <?php endif; ?>

    <form action="index.php" method="post" enctype="multipart/form-data" id="convert-form">
        <div class="form-group">
            <label for="srt_file">1. Select SRT File</label>
            <input type="file" id="srt_file" name="srt_file" accept=".srt" required>
        </div>

        <div class="form-group">
            <label for="project_name">2. Project Name (for DOCX title)</label>
            <input type="text" id="project_name" name="project_name" placeholder="Will be inferred from filename if left blank">
        </div>

        <div class="form-group">
            <label>3. Choose Output Format</label>
            <div class="radio-group">
                <label>
                    <input type="radio" name="output_format" value="csv" checked> CSV (for Excel)
                </label>
                <label>
                    <input type="radio" name="output_format" value="docx" <?php if (!$phpWordInstalled) echo 'disabled'; ?>> DOCX (for Word)
                </label>
            </div>
        </div>

        <button type="submit" class="btn">Convert and Download</button>
    </form>
</div>

<footer>
    Idea by Pilar Terán &bull; Program by @edfasano70
</footer>

<script>
    // Script to auto-fill project name from the selected filename
    document.getElementById('srt_file').addEventListener('change', function() {
        const fileInput = this;
        const projectNameInput = document.getElementById('project_name');
        if (fileInput.files.length > 0) {
            let filename = fileInput.files[0].name;
            // Remove the .srt extension
            let basename = filename.substring(0, filename.lastIndexOf('.')) || filename;
            projectNameInput.value = basename;
        }
    });
</script>

</body>
</html>