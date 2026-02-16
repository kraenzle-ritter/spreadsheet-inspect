<?php

namespace App\Commands;

use Dompdf\Dompdf;
use Dompdf\Options;
use Illuminate\Support\Collection;
use LaravelZero\Framework\Commands\Command;
use PhpOffice\PhpSpreadsheet\IOFactory;
use Rap2hpoutre\FastExcel\FastExcel;

class SpreadsheetCommand extends Command
{
    protected $signature = 'spreadsheet
                            {file : Path to the Excel file}
                            {--sheets : List all sheet names}
                            {--sheet= : Inspect a specific sheet by name or index}
                            {--column= : Cross search for values from this column}
                            {--cross-sheet= : Only check this sheet}
                            {--target-column= : Only compare against this column in target sheets}
                            {--debug : Show matching values side by side for debugging}
                            {--images : Count and list images/thumbnails in the sheet}
                            {--extract-images= : Extract images to this directory}
                            {--output= : Output format: console (default), html, pdf}
                            {--output-file= : Output file path (required for html/pdf)}
                            {--memory=2000 : Memory limit in MB}';

    protected array $reportData = [];

    protected $description = 'Inspect spreadsheet files (Excel, LibreOffice) - analyze sheets, columns, images';

    public function handle(): int
    {
        $this->setMemoryLimit();

        $file = $this->argument('file');

        // Expand ~ to home directory
        if (str_starts_with($file, '~')) {
            $file = getenv('HOME') . substr($file, 1);
        }

        if (!file_exists($file)) {
            $this->error("File not found: $file");
            return 1;
        }

        $outputFormat = $this->option('output') ?? 'console';
        $outputFile = $this->option('output-file');

        if (in_array($outputFormat, ['html', 'pdf']) && !$outputFile) {
            $this->error("--output-file is required when using --output=$outputFormat");
            return 1;
        }

        $this->reportData = [
            'file' => basename($file),
            'filePath' => $file,
            'generatedAt' => date('Y-m-d H:i:s'),
            'sheets' => [],
            'analysis' => null,
            'images' => null,
            'crossSheet' => null,
        ];

        $sheetNames = $this->getSheetNames($file);
        $this->reportData['sheets'] = $sheetNames;

        if ($outputFormat === 'console') {
            $this->line("## Available sheets\n");
            foreach ($sheetNames as $i => $s) {
                $this->line("- **[$i]** `$s`");
            }
            $this->line("");
        }

        // --sheets: Nur Sheetnamen anzeigen
        if ($this->option('sheets')) {
            return $this->outputReport($outputFormat, $outputFile);
        }

        $sheetOption = $this->option('sheet');
        $sheetNumber = $this->getSheetNumber($sheetNames, $sheetOption);

        if (!$sheetNumber) {
            $this->error("Sheet $sheetOption not found.");
            return 1;
        }

        $sheetName = $sheetNames[$sheetNumber] ?? "Sheet $sheetNumber";
        $this->reportData['selectedSheet'] = ['index' => $sheetNumber, 'name' => $sheetName];

        $rows = collect((new FastExcel())->sheet($sheetNumber)->import($file));
        $rows = $this->sanitizeSheet($rows);

        if ($rows->isEmpty()) {
            $this->warn("No data found in sheet '$sheetName' ($sheetNumber)");
            return $this->outputReport($outputFormat, $outputFile);
        }

        if ($outputFormat === 'console') {
            $this->info('');
            $this->info("# Sheet `$sheetName` (Index: $sheetNumber)\n");
        }

        if ($this->option('sheet') && !$this->option('column')) {
            $this->analyzeSheetData($rows, $outputFormat === 'console');
        }

        // Handle --images and --extract-images options
        if ($this->option('images') || $this->option('extract-images')) {
            $this->analyzeImages($file, $sheetNumber, $sheetNames, $outputFormat === 'console');
        }

        if ($this->option('column')) {
            $column = $this->option('column');
            $this->analyzeCrossSheetUsage($file, $sheetNumber, $column, $sheetNames, $outputFormat === 'console');
        }

        return $this->outputReport($outputFormat, $outputFile);
    }

    protected function outputReport(string $format, ?string $outputFile): int
    {
        if ($format === 'console') {
            return 0;
        }

        $html = $this->generateHtml();

        if ($format === 'html') {
            // Expand ~ to home directory
            if (str_starts_with($outputFile, '~')) {
                $outputFile = getenv('HOME') . substr($outputFile, 1);
            }
            file_put_contents($outputFile, $html);
            $this->info("HTML report saved to: $outputFile");
            return 0;
        }

        if ($format === 'pdf') {
            // Expand ~ to home directory
            if (str_starts_with($outputFile, '~')) {
                $outputFile = getenv('HOME') . substr($outputFile, 1);
            }
            $this->generatePdf($html, $outputFile);
            $this->info("PDF report saved to: $outputFile");
            return 0;
        }

        $this->error("Unknown output format: $format");
        return 1;
    }

    protected function generateHtml(): string
    {
        $data = $this->reportData;
        $logoDataUri = $this->getInvertedLogoDataUri();
        $html = <<<HTML
<!DOCTYPE html>
<html lang="en">
<head>
    <meta charset="UTF-8">
    <meta name="viewport" content="width=device-width, initial-scale=1.0">
    <title>Spreadsheet Report: {$data['file']}</title>
    <style>
        * { box-sizing: border-box; }
        body {
            font-family: -apple-system, BlinkMacSystemFont, 'Segoe UI', Roboto, Oxygen, Ubuntu, sans-serif;
            line-height: 1.6;
            max-width: 1000px;
            margin: 0 auto;
            padding: 20px;
            padding-top: 10px;
            color: #333;
            position: relative;
        }
        .logo { width: 60px; height: 60px; position: absolute; top: 10px; right: 20px; }
        h1 { color: #6b7280; border-bottom: 3px solid #9ca3af; padding-bottom: 10px; margin-top: 0; }
        h2 { color: #6b7280; margin-top: 30px; border-bottom: 1px solid #d1d5db; padding-bottom: 5px; }
        h3 { color: #9ca3af; margin-top: 20px; }
        .meta { color: #6b7280; font-size: 0.9em; margin-bottom: 20px; }
        .stat { background: #f3f4f6; padding: 3px 8px; border-radius: 4px; font-family: monospace; }
        .column-card {
            background: #f9fafb;
            border: 1px solid #e5e7eb;
            border-radius: 8px;
            padding: 15px;
            margin: 15px 0;
        }
        .column-name { font-weight: bold; color: #374151; font-size: 1.1em; }
        .values-list {
            background: white;
            border: 1px solid #e5e7eb;
            border-radius: 4px;
            padding: 10px;
            margin-top: 10px;
        }
        .value-item { padding: 3px 0; border-bottom: 1px solid #f3f4f6; }
        .value-item:last-child { border-bottom: none; }
        .count { color: #6b7280; font-size: 0.9em; }
        .progress-bar {
            background: #e5e7eb;
            border-radius: 4px;
            height: 8px;
            margin-top: 5px;
        }
        .progress-fill {
            background: #22c55e;
            height: 100%;
            border-radius: 4px;
        }
        .sheet-list { list-style: none; padding: 0; }
        .sheet-list li {
            padding: 8px 12px;
            background: #f9fafb;
            margin: 5px 0;
            border-radius: 4px;
            border-left: 4px solid #3b82f6;
        }
        .sheet-list .index { color: #6b7280; font-weight: bold; }
        table { width: 100%; border-collapse: collapse; margin: 15px 0; }
        th, td { padding: 10px; text-align: left; border-bottom: 1px solid #e5e7eb; }
        th { background: #f9fafb; font-weight: 600; }
        .image-stats { display: flex; gap: 20px; flex-wrap: wrap; }
        .image-stat { background: #dbeafe; padding: 10px 15px; border-radius: 8px; }
        footer { margin-top: 40px; padding-top: 20px; border-top: 1px solid #e5e7eb; color: #9ca3af; font-size: 0.85em; }
    </style>
</head>
<body>
    <img src="{$logoDataUri}" alt="Logo" class="logo">
    <h1>Spreadsheet Report</h1>
    <div class="meta">
        <strong>File:</strong> {$data['file']}<br>
        <strong>Generated:</strong> {$data['generatedAt']}
    </div>

    <h2>Available Sheets</h2>
    <ul class="sheet-list">
HTML;

        foreach ($data['sheets'] as $index => $name) {
            $selected = isset($data['selectedSheet']) && $data['selectedSheet']['index'] === $index ? ' style="border-left-color: #28a745;"' : '';
            $html .= "<li{$selected}><span class=\"index\">[$index]</span> $name</li>";
        }

        $html .= "</ul>";

        if (isset($data['analysis'])) {
            $analysis = $data['analysis'];
            $html .= "<h2>Sheet Analysis: {$data['selectedSheet']['name']}</h2>";
            $html .= "<p><strong>Total Rows:</strong> <span class=\"stat\">{$analysis['totalRows']}</span></p>";

            foreach ($analysis['columns'] as $col) {
                $html .= "<div class=\"column-card\">";
                $html .= "<div class=\"column-name\">{$col['name']}</div>";
                $html .= "<p><strong>Filled:</strong> {$col['filled']} / {$col['total']} ({$col['percent']}%)</p>";
                $html .= "<div class=\"progress-bar\"><div class=\"progress-fill\" style=\"width: {$col['percent']}%\"></div></div>";
                $html .= "<p><strong>Distinct Values:</strong> {$col['distinctCount']}</p>";

                if (!empty($col['values'])) {
                    $html .= "<div class=\"values-list\">";
                    foreach ($col['values'] as $val => $count) {
                        $escapedVal = htmlspecialchars((string) $val);
                        $html .= "<div class=\"value-item\"><code>$escapedVal</code> <span class=\"count\">($count)</span></div>";
                    }
                    if ($col['hasMore'] ?? false) {
                        $html .= "<div class=\"value-item\"><em>… and " . ($col['distinctCount'] - 10) . " more</em></div>";
                    }
                    $html .= "</div>";
                }
                $html .= "</div>";
            }
        }

        if (isset($data['images'])) {
            $images = $data['images'];
            $html .= "<h2>Images in Sheet</h2>";
            $html .= "<p><strong>Total Images:</strong> <span class=\"stat\">{$images['total']}</span></p>";

            if ($images['total'] > 0) {
                $html .= "<div class=\"image-stats\">";
                foreach ($images['byColumn'] as $col => $count) {
                    $html .= "<div class=\"image-stat\"><strong>Column $col:</strong> $count image(s)</div>";
                }
                $html .= "</div>";
                $html .= "<p><strong>Rows with images:</strong> {$images['rowsWithImages']}</p>";
            }
        }

        if (isset($data['crossSheet'])) {
            $cross = $data['crossSheet'];
            $html .= "<h2>Cross-Sheet References</h2>";
            $html .= "<p><strong>Source Column:</strong> {$cross['sourceColumn']}</p>";
            $html .= "<p><strong>Values Found:</strong> {$cross['found']} / {$cross['total']} ({$cross['percent']}%)</p>";

            if (!empty($cross['matches'])) {
                $html .= "<table><thead><tr><th>Sheet</th><th>Column</th><th>Matches</th></tr></thead><tbody>";
                foreach ($cross['matches'] as $match) {
                    $html .= "<tr><td>{$match['sheetName']}</td><td>{$match['column']}</td><td>{$match['count']}</td></tr>";
                }
                $html .= "</tbody></table>";
            }
        }

        $html .= <<<HTML
    <footer>
        Generated by <strong>Spreadsheet Inspect</strong> • {$data['generatedAt']}<br>
        <a href="https://github.com/kraenzle-ritter/spreadsheet-inspect">https://github.com/kraenzle-ritter/spreadsheet-inspect</a>
    </footer>
</body>
</html>
HTML;

        return $html;
    }

    protected function generatePdf(string $html, string $outputFile): void
    {
        $options = new Options();
        $options->set('isHtml5ParserEnabled', true);
        $options->set('isRemoteEnabled', true);
        $options->set('defaultFont', 'DejaVu Sans');

        $dompdf = new Dompdf($options);
        $dompdf->loadHtml($html);
        $dompdf->setPaper('A4', 'portrait');
        $dompdf->render();

        file_put_contents($outputFile, $dompdf->output());
    }

    protected function setMemoryLimit(): void
    {
        $memory = $this->option('memory');
        if (is_numeric($memory) && $memory > 0) {
            ini_set('memory_limit', $memory . 'M');
        } else {
            $this->warn("Invalid memory value provided. Skipping memory_limit change.");
        }
    }

    protected function getSheetNames(string $file): array
    {
        $reader = IOFactory::createReaderForFile($file);
        $reader->setReadDataOnly(true);
        $sheets = $reader->listWorksheetNames($file);

        $result = [];
        $i = 1;
        foreach ($sheets as $sheet) {
            $result[$i++] = $sheet;
        }
        return $result;
    }

    protected function getSheetNumber(array $sheetNames, ?string $sheetOption): ?int
    {
        if ($sheetOption === null) {
            $this->warn("No sheet specified. Use --sheet=1 or --sheet=SheetName");
            return null;
        }

        // Ist sheet eine Zahl?
        if (is_numeric($sheetOption)) {
            $index = (int) $sheetOption;
            if (!isset($sheetNames[$index])) {
                $this->error("Sheet index '$index' is out of range. Max index: " . count($sheetNames));
                return null;
            }
            return $index;
        }

        // Ist sheet ein Name?
        if (in_array($sheetOption, $sheetNames, true)) {
            return array_search($sheetOption, $sheetNames, true);
        }

        $this->error("Sheet '$sheetOption' not found in list of sheets.");
        return null;
    }

    protected function analyzeSheetData(Collection $rows, bool $consoleOutput = true): void
    {
        $headers = array_keys($rows->first());
        $totalRows = $rows->count();

        $this->reportData['analysis'] = [
            'totalRows' => $totalRows,
            'columns' => [],
        ];

        if ($consoleOutput) {
            $this->info("\n## Sheet statistics\n");
            $this->line("- **Rows** (excluding header): `$totalRows`\n");
        }

        foreach ($headers as $header) {
            $values = $rows->map(fn($row) => $row[$header] ?? null);

            $nonEmpty = $values->filter(fn($v) => $v !== null && $v !== '');

            $count = $nonEmpty->count();
            $percent = $totalRows > 0 ? round(($count / $totalRows) * 100, 2) : 0;

            $distinct = $nonEmpty->countBy()->sortDesc();
            $distinctCount = $distinct->count();

            $columnData = [
                'name' => $header,
                'filled' => $count,
                'total' => $totalRows,
                'percent' => $percent,
                'distinctCount' => $distinctCount,
                'values' => [],
                'hasMore' => false,
            ];

            if ($distinctCount > 0 && $distinctCount <= 20) {
                $columnData['values'] = $distinct->toArray();
            } elseif ($distinctCount > 20) {
                $columnData['values'] = $distinct->take(10)->toArray();
                $columnData['hasMore'] = true;
            }

            $this->reportData['analysis']['columns'][] = $columnData;

            if ($consoleOutput) {
                $this->line("### `$header`\n");

                if ($count === 0 && stripos($header, 'bild') !== false) {
                    $this->line("- **Filled**: `$count / $totalRows` ($percent%) *Images may be embedded as drawings (use --images)*");
                } else {
                    $this->line("- **Filled**: `$count / $totalRows` ($percent%)");
                }

                $this->line("- **Distinct**: `$distinctCount`");

                if ($distinctCount > 0 && $distinctCount <= 20) {
                    $this->line("\n  Values:");
                    foreach ($distinct as $val => $occurrences) {
                        $this->line("  - `" . $this->truncateValue($val, 100) . "` ($occurrences)");
                    }
                } elseif ($distinctCount > 20) {
                    $this->line("\n  Top 10 (of $distinctCount):");
                    foreach ($distinct->take(10) as $val => $occurrences) {
                        $this->line("  - `" . $this->truncateValue($val, 100) . "` ($occurrences)");
                    }
                    $this->line("  - *… and " . ($distinctCount - 10) . " more*");
                }
                $this->line("");
            }
        }
    }

    protected function truncateValue($value, int $maxLength = 100): string
    {
        $str = (string) $value;
        if (strlen($str) > $maxLength) {
            return substr($str, 0, $maxLength) . '…';
        }
        return $str;
    }

    protected function sanitizeSheet(Collection $rows): Collection
    {
        return $rows->map(fn($row) =>
            collect($row)->map(fn($value) =>
                $value instanceof \DateTimeInterface ? $value->format('Y-m-d') : $value
            )->toArray()
        );
    }

    protected function analyzeCrossSheetUsage(string $file, int $sourceSheet, string $sourceColumn, array $sheetNames, bool $consoleOutput = true): void
    {
        $sourceSheetName = $sheetNames[$sourceSheet] ?? "Unknown";

        if ($consoleOutput) {
            $this->info("\nCross-sheet reference check for column '$sourceColumn' in sheet $sourceSheetName:");
        }

        $sourceValues = $this->loadUniqueValuesFromColumn($file, $sourceSheet, $sourceColumn);
        $sourceSet = $sourceValues->all();

        $targetSheetNumber = $this->resolveCrossSheetTarget($sheetNames);
        if ($targetSheetNumber === false) return;

        $targetColumn = $this->option('target-column');
        $matches = $this->findCrossSheetMatches($file, $sourceSet, $targetColumn, $sheetNames, $sourceSheet, $targetSheetNumber, $consoleOutput);

        $this->outputCrossSheetSummary($matches, $sourceValues, $sheetNames, $sourceColumn, $consoleOutput);
    }

    protected function loadUniqueValuesFromColumn(string $file, int $sheetIndex, string $column): Collection
    {
        $rows = collect((new FastExcel)->sheet($sheetIndex)->import($file));

        return $rows->map(fn($row) => $row[$column] ?? null)
            ->map(fn($v) => $v instanceof \DateTimeInterface ? $v->format('Y-m-d') : $v)
            ->filter(fn($v) => $v !== null && $v !== '')
            ->values();
    }

    protected function resolveCrossSheetTarget(array $sheetNames): int|false|null
    {
        $option = $this->option('cross-sheet');
        if (!$option) return null;

        $target = $this->getSheetNumber($sheetNames, $option);
        if ($target === null) {
            $this->error("Cross-sheet '$option' not found.");
            return false;
        }

        return $target;
    }

    protected function findCrossSheetMatches(
        string $file,
        array $sourceSet,
        ?string $column,
        array $sheetNames,
        int $sourceSheet,
        ?int $onlySheet = null,
        bool $consoleOutput = true
    ): array {
        $matches = [];

        foreach ($sheetNames as $index => $name) {
            if ($index === $sourceSheet) continue;
            if ($onlySheet !== null && $index !== $onlySheet) continue;

            if ($consoleOutput) {
                $this->line("Checking sheet: $name ($index)");
            }
            $rows = collect((new FastExcel)->sheet($index)->import($file));

            if ($this->option('debug')) {
                $rows = $rows->take(100);
                if ($consoleOutput) {
                    $this->warn("Debug mode: only checking first 100 rows");
                }
            }

            if ($rows->isEmpty()) continue;

            $headers = array_keys($rows->first());
            if ($column && !in_array($column, $headers)) {
                if ($consoleOutput) {
                    $this->warn("Column '$column' not found in '$name'");
                }
                continue;
            }

            $values = $rows->map(fn($row) => $row[$column] ?? null)
                ->map(fn($v) => $v instanceof \DateTimeInterface ? $v->format('Y-m-d') : $v)
                ->filter();

            $intersected = [];
            foreach ($sourceSet as $v) {
                if (in_array($v, $values->toArray())) {
                    $intersected[] = $v;
                }
            }

            if (!empty($intersected)) {
                $matches[] = [
                    'sheet' => $index,
                    'sheetName' => $name,
                    'column' => $column,
                    'count' => count($intersected),
                    'values' => collect($intersected)->unique()->values(),
                ];
            }
        }

        return $matches;
    }

    protected function outputCrossSheetSummary(array $matches, Collection $sourceValues, array $sheetNames, string $sourceColumn, bool $consoleOutput = true): void
    {
        $total = $sourceValues->count();
        $found = collect($matches)->sum('count');
        $percent = $total > 0 ? round(($found / $total) * 100, 2) : 0;

        $this->reportData['crossSheet'] = [
            'sourceColumn' => $sourceColumn,
            'total' => $total,
            'found' => $found,
            'percent' => $percent,
            'matches' => $matches,
        ];

        if ($consoleOutput) {
            $this->line("Values found: $found / $total ($percent%)");

            foreach ($matches as $match) {
                $sheetName = $sheetNames[$match['sheet']] ?? "Unknown";
                $this->line(" - In sheet $sheetName ({$match['sheet']}), column '{$match['column']}' → {$match['count']} match(es)");
                if ($match['values']->count() <= 10) {
                    $this->line("   → " . $match['values']->implode(', '));
                }
            }
        }
    }

    protected function analyzeImages(string $file, int $sheetNumber, array $sheetNames, bool $consoleOutput = true): void
    {
        $sheetName = $sheetNames[$sheetNumber] ?? "Unknown";

        if ($consoleOutput) {
            $this->info("\n## Images in sheet `$sheetName`\n");
        }

        $reader = IOFactory::createReaderForFile($file);
        $spreadsheet = $reader->load($file);
        $worksheet = $spreadsheet->getSheet($sheetNumber - 1);

        $drawings = $worksheet->getDrawingCollection();
        $imageCount = count($drawings);

        $imagesByColumn = [];
        $imagesByRow = [];

        foreach ($drawings as $drawing) {
            $coordinates = $drawing->getCoordinates();
            preg_match('/^([A-Z]+)(\d+)$/', $coordinates, $matches);
            $column = $matches[1] ?? 'Unknown';
            $row = $matches[2] ?? 'Unknown';

            $imagesByColumn[$column] = ($imagesByColumn[$column] ?? 0) + 1;
            $imagesByRow[$row] = ($imagesByRow[$row] ?? 0) + 1;
        }

        $this->reportData['images'] = [
            'total' => $imageCount,
            'byColumn' => $imagesByColumn,
            'rowsWithImages' => count($imagesByRow),
        ];

        if ($consoleOutput) {
            $this->line("- **Total images found**: `$imageCount`");

            if ($imageCount === 0) {
                $this->warn("No images found in this sheet.");
                return;
            }

            $this->line("\n### Images by column\n");
            foreach ($imagesByColumn as $col => $count) {
                $this->line("- **Column $col**: `$count` image(s)");
            }

            $this->line("\n### Distribution\n");
            $this->line("- **Rows with images**: `" . count($imagesByRow) . "`");
        }

        $extractDir = $this->option('extract-images');
        if ($extractDir) {
            $this->extractImages($drawings, $extractDir, $sheetName);
        }

        if ($consoleOutput && $this->option('debug') && $imageCount > 0) {
            $this->line("\nSample image details (first 10):");
            $i = 0;
            foreach ($drawings as $drawing) {
                if ($i >= 10) break;
                $this->line("   [{$drawing->getCoordinates()}] " .
                    "Name: " . ($drawing->getName() ?: 'unnamed') . ", " .
                    "Description: " . ($drawing->getDescription() ?: 'none'));
                $i++;
            }
        }
    }

    protected function extractImages($drawings, string $outputDir, string $sheetName): void
    {
        if (str_starts_with($outputDir, '~')) {
            $outputDir = getenv('HOME') . substr($outputDir, 1);
        }

        if (!is_dir($outputDir)) {
            mkdir($outputDir, 0755, true);
            $this->line("Created directory: $outputDir");
        }

        $extracted = 0;
        $failed = 0;

        foreach ($drawings as $index => $drawing) {
            try {
                $coordinates = $drawing->getCoordinates();

                if ($drawing instanceof \PhpOffice\PhpSpreadsheet\Worksheet\Drawing) {
                    $path = $drawing->getPath();
                    $extension = pathinfo($path, PATHINFO_EXTENSION) ?: 'png';
                    $filename = sprintf('%s_%s_%03d.%s',
                        preg_replace('/[^a-zA-Z0-9]/', '_', $sheetName),
                        $coordinates,
                        $index + 1,
                        $extension
                    );

                    $contents = file_get_contents($path);
                    if ($contents !== false) {
                        file_put_contents("$outputDir/$filename", $contents);
                        $extracted++;
                    } else {
                        $failed++;
                    }
                } elseif ($drawing instanceof \PhpOffice\PhpSpreadsheet\Worksheet\MemoryDrawing) {
                    $extension = $this->getMimeTypeExtension($drawing->getMimeType());
                    $filename = sprintf('%s_%s_%03d.%s',
                        preg_replace('/[^a-zA-Z0-9]/', '_', $sheetName),
                        $coordinates,
                        $index + 1,
                        $extension
                    );

                    ob_start();
                    $renderFunction = $drawing->getRenderingFunction();
                    call_user_func($renderFunction, $drawing->getImageResource());
                    $imageData = ob_get_clean();

                    if ($imageData) {
                        file_put_contents("$outputDir/$filename", $imageData);
                        $extracted++;
                    } else {
                        $failed++;
                    }
                }
            } catch (\Throwable $e) {
                $failed++;
                if ($this->option('debug')) {
                    $this->warn("Failed to extract image $index: " . $e->getMessage());
                }
            }
        }

        $this->info("\nExtracted $extracted image(s) to: $outputDir");
        if ($failed > 0) {
            $this->warn("Failed to extract $failed image(s)");
        }
    }

    protected function getMimeTypeExtension(string $mimeType): string
    {
        return match ($mimeType) {
            \PhpOffice\PhpSpreadsheet\Worksheet\MemoryDrawing::MIMETYPE_PNG => 'png',
            \PhpOffice\PhpSpreadsheet\Worksheet\MemoryDrawing::MIMETYPE_GIF => 'gif',
            \PhpOffice\PhpSpreadsheet\Worksheet\MemoryDrawing::MIMETYPE_JPEG => 'jpg',
            default => 'png',
        };
    }

    protected function getInvertedLogoDataUri(): string
    {
        $logoUrl = 'https://avatars.githubusercontent.com/u/67142673?s=400&u=534a1480b999c454e525885361ee7aeb541b230a&v=4';
        
        try {
            $imageData = @file_get_contents($logoUrl);
            if ($imageData === false) {
                return '';
            }

            $image = @imagecreatefromstring($imageData);
            if ($image === false) {
                return '';
            }

            // Invert colors
            imagefilter($image, IMG_FILTER_NEGATE);
            // Convert to grayscale
            imagefilter($image, IMG_FILTER_GRAYSCALE);

            // Capture as PNG
            ob_start();
            imagepng($image);
            $invertedData = ob_get_clean();
            imagedestroy($image);

            return 'data:image/png;base64,' . base64_encode($invertedData);
        } catch (\Throwable $e) {
            return '';
        }
    }
}
