<?php

namespace App\Commands;

use Illuminate\Support\Collection;
use LaravelZero\Framework\Commands\Command;
use PhpOffice\PhpSpreadsheet\IOFactory;
use Rap2hpoutre\FastExcel\FastExcel;

class InspectCommand extends Command
{
    protected $signature = 'inspect
                            {file : Path to the Excel file}
                            {--sheets : List all sheet names}
                            {--sheet= : Inspect a specific sheet by name or index}
                            {--column= : Cross search for values from this column}
                            {--cross-sheet= : Only check this sheet}
                            {--target-column= : Only compare against this column in target sheets}
                            {--debug : Show matching values side by side for debugging}
                            {--images : Count and list images/thumbnails in the sheet}
                            {--extract-images= : Extract images to this directory}
                            {--memory=2000 : Memory limit in MB}';

    protected $description = 'Inspect Excel file (list sheets, column headers, unique values, images)';

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

        $sheetNames = $this->getSheetNames($file);

        $this->line("## Available sheets\n");
        foreach ($sheetNames as $i => $s) {
            $this->line("- **[$i]** `$s`");
        }
        $this->line("");

        // --sheets: Nur Sheetnamen anzeigen
        if ($this->option('sheets')) {
            return 0;
        }

        $sheetOption = $this->option('sheet');
        $sheetNumber = $this->getSheetNumber($sheetNames, $sheetOption);

        if (!$sheetNumber) {
            $this->error("Sheet $sheetOption not found.");
            return 1;
        }

        $sheetName = $sheetNames[$sheetNumber] ?? "Sheet $sheetNumber";

        $rows = collect((new FastExcel())->sheet($sheetNumber)->import($file));
        $rows = $this->sanitizeSheet($rows);

        if ($rows->isEmpty()) {
            $this->warn("No data found in sheet '$sheetName' ($sheetNumber)");
            return 0;
        }

        $this->info('');
        $this->info("# Sheet `$sheetName` (Index: $sheetNumber)\n");

        if ($this->option('sheet') && !$this->option('column')) {
            $this->analyzeSheetData($rows);
        }

        // Handle --images and --extract-images options
        if ($this->option('images') || $this->option('extract-images')) {
            $this->analyzeImages($file, $sheetNumber, $sheetNames);
        }

        if ($this->option('column')) {
            $column = $this->option('column');
            $this->analyzeCrossSheetUsage($file, $sheetNumber, $column, $sheetNames);
        }

        return 0;
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

    protected function analyzeSheetData(Collection $rows): void
    {
        $headers = array_keys($rows->first());
        $totalRows = $rows->count();

        $this->info("\n## Sheet statistics\n");
        $this->line("- **Rows** (excluding header): `$totalRows`\n");

        foreach ($headers as $header) {
            $values = $rows->map(fn($row) => $row[$header] ?? null);

            $nonEmpty = $values->filter(fn($v) => $v !== null && $v !== '');

            $count = $nonEmpty->count();
            $percent = $totalRows > 0 ? round(($count / $totalRows) * 100, 2) : 0;

            $this->line("### `$header`\n");

            if ($count === 0 && stripos($header, 'bild') !== false) {
                $this->line("- **Filled**: `$count / $totalRows` ($percent%) *Images may be embedded as drawings (use --images)*");
            } else {
                $this->line("- **Filled**: `$count / $totalRows` ($percent%)");
            }

            $distinct = $nonEmpty->countBy()->sortDesc();
            $distinctCount = $distinct->count();
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

    protected function analyzeCrossSheetUsage(string $file, int $sourceSheet, string $sourceColumn, array $sheetNames): void
    {
        $sourceSheetName = $sheetNames[$sourceSheet] ?? "Unknown";
        $this->info("\nCross-sheet reference check for column '$sourceColumn' in sheet $sourceSheetName:");

        $sourceValues = $this->loadUniqueValuesFromColumn($file, $sourceSheet, $sourceColumn);
        $sourceSet = $sourceValues->all();

        $targetSheetNumber = $this->resolveCrossSheetTarget($sheetNames);
        if ($targetSheetNumber === false) return;

        $targetColumn = $this->option('target-column');
        $matches = $this->findCrossSheetMatches($file, $sourceSet, $targetColumn, $sheetNames, $sourceSheet, $targetSheetNumber);

        $this->outputCrossSheetSummary($matches, $sourceValues, $sheetNames);
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
        ?int $onlySheet = null
    ): array {
        $matches = [];

        foreach ($sheetNames as $index => $name) {
            if ($index === $sourceSheet) continue;
            if ($onlySheet !== null && $index !== $onlySheet) continue;

            $this->line("Checking sheet: $name ($index)");
            $rows = collect((new FastExcel)->sheet($index)->import($file));

            if ($this->option('debug')) {
                $rows = $rows->take(100);
                $this->warn("Debug mode: only checking first 100 rows");
            }

            if ($rows->isEmpty()) continue;

            $headers = array_keys($rows->first());
            if ($column && !in_array($column, $headers)) {
                $this->warn("Column '$column' not found in '$name'");
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
                    'column' => $column,
                    'count' => count($intersected),
                    'values' => collect($intersected)->unique()->values(),
                ];
            }
        }

        return $matches;
    }

    protected function outputCrossSheetSummary(array $matches, Collection $sourceValues, array $sheetNames): void
    {
        $total = $sourceValues->count();
        $found = collect($matches)->sum('count');
        $percent = $total > 0 ? round(($found / $total) * 100, 2) : 0;

        $this->line("Values found: $found / $total ($percent%)");

        foreach ($matches as $match) {
            $sheetName = $sheetNames[$match['sheet']] ?? "Unknown";
            $this->line(" - In sheet $sheetName ({$match['sheet']}), column '{$match['column']}' → {$match['count']} match(es)");
            if ($match['values']->count() <= 10) {
                $this->line("   → " . $match['values']->implode(', '));
            }
        }
    }

    protected function analyzeImages(string $file, int $sheetNumber, array $sheetNames): void
    {
        $sheetName = $sheetNames[$sheetNumber] ?? "Unknown";
        $this->info("\n## Images in sheet `$sheetName`\n");

        $reader = IOFactory::createReaderForFile($file);
        $spreadsheet = $reader->load($file);
        $worksheet = $spreadsheet->getSheet($sheetNumber - 1);

        $drawings = $worksheet->getDrawingCollection();
        $imageCount = count($drawings);

        $this->line("- **Total images found**: `$imageCount`");

        if ($imageCount === 0) {
            $this->warn("No images found in this sheet.");
            return;
        }

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

        $this->line("\n### Images by column\n");
        foreach ($imagesByColumn as $col => $count) {
            $this->line("- **Column $col**: `$count` image(s)");
        }

        $this->line("\n### Distribution\n");
        $this->line("- **Rows with images**: `" . count($imagesByRow) . "`");

        $extractDir = $this->option('extract-images');
        if ($extractDir) {
            $this->extractImages($drawings, $extractDir, $sheetName);
        }

        if ($this->option('debug') && $imageCount > 0) {
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
}
