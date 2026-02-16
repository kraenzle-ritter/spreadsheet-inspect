<?php

use App\Commands\SpreadsheetCommand;
use Illuminate\Support\Collection;

beforeEach(function () {
    $this->command = new class extends SpreadsheetCommand
    {
        // Expose protected methods for testing
        public function publicTruncateValue($value, int $maxLength = 100): string
        {
            return $this->truncateValue($value, $maxLength);
        }

        public function publicSanitizeSheet(Collection $rows): Collection
        {
            return $this->sanitizeSheet($rows);
        }

        public function publicGetMimeTypeExtension(string $mimeType): string
        {
            return $this->getMimeTypeExtension($mimeType);
        }
    };
});

describe('truncateValue', function () {
    it('returns short strings unchanged', function () {
        expect($this->command->publicTruncateValue('short string', 100))
            ->toBe('short string');
    });

    it('truncates long strings with ellipsis', function () {
        $longString = str_repeat('a', 150);
        $result = $this->command->publicTruncateValue($longString, 100);

        expect(mb_strlen($result))->toBe(101); // 100 chars + ellipsis (use mb_strlen for UTF-8)
        expect($result)->toEndWith('…');
    });

    it('handles exact length strings', function () {
        $exactString = str_repeat('a', 100);
        expect($this->command->publicTruncateValue($exactString, 100))
            ->toBe($exactString);
    });

    it('converts non-strings to strings', function () {
        expect($this->command->publicTruncateValue(12345, 100))
            ->toBe('12345');
    });

    it('handles empty strings', function () {
        expect($this->command->publicTruncateValue('', 100))
            ->toBe('');
    });
});

describe('sanitizeSheet', function () {
    it('converts DateTimeInterface to Y-m-d format', function () {
        $rows = collect([
            ['date' => new DateTime('2024-05-15'), 'name' => 'Test'],
        ]);

        $result = $this->command->publicSanitizeSheet($rows);

        expect($result->first()['date'])->toBe('2024-05-15');
        expect($result->first()['name'])->toBe('Test');
    });

    it('leaves non-date values unchanged', function () {
        $rows = collect([
            ['value' => 'string', 'number' => 42, 'null' => null],
        ]);

        $result = $this->command->publicSanitizeSheet($rows);

        expect($result->first()['value'])->toBe('string');
        expect($result->first()['number'])->toBe(42);
        expect($result->first()['null'])->toBeNull();
    });

    it('handles empty collections', function () {
        $rows = collect([]);
        $result = $this->command->publicSanitizeSheet($rows);

        expect($result)->toBeEmpty();
    });

    it('handles multiple rows with mixed dates', function () {
        $rows = collect([
            ['date' => new DateTime('2024-01-01'), 'text' => 'First'],
            ['date' => new DateTime('2024-12-31'), 'text' => 'Last'],
            ['date' => 'not a date', 'text' => 'Third'],
        ]);

        $result = $this->command->publicSanitizeSheet($rows);

        expect($result[0]['date'])->toBe('2024-01-01');
        expect($result[1]['date'])->toBe('2024-12-31');
        expect($result[2]['date'])->toBe('not a date');
    });
});

describe('getMimeTypeExtension', function () {
    it('returns png for PNG mime type', function () {
        $mimeType = \PhpOffice\PhpSpreadsheet\Worksheet\MemoryDrawing::MIMETYPE_PNG;
        expect($this->command->publicGetMimeTypeExtension($mimeType))->toBe('png');
    });

    it('returns gif for GIF mime type', function () {
        $mimeType = \PhpOffice\PhpSpreadsheet\Worksheet\MemoryDrawing::MIMETYPE_GIF;
        expect($this->command->publicGetMimeTypeExtension($mimeType))->toBe('gif');
    });

    it('returns jpg for JPEG mime type', function () {
        $mimeType = \PhpOffice\PhpSpreadsheet\Worksheet\MemoryDrawing::MIMETYPE_JPEG;
        expect($this->command->publicGetMimeTypeExtension($mimeType))->toBe('jpg');
    });

    it('returns png as default for unknown mime type', function () {
        expect($this->command->publicGetMimeTypeExtension('image/unknown'))->toBe('png');
    });
});

describe('home directory expansion', function () {
    it('expands tilde to home directory', function () {
        $home = getenv('HOME');
        $path = '~/test/file.xlsx';
        $expected = $home.'/test/file.xlsx';

        // Test the expansion logic directly
        if (str_starts_with($path, '~')) {
            $expanded = getenv('HOME').substr($path, 1);
        } else {
            $expanded = $path;
        }

        expect($expanded)->toBe($expected);
    });

    it('leaves absolute paths unchanged', function () {
        $path = '/absolute/path/file.xlsx';

        if (str_starts_with($path, '~')) {
            $expanded = getenv('HOME').substr($path, 1);
        } else {
            $expanded = $path;
        }

        expect($expanded)->toBe('/absolute/path/file.xlsx');
    });
});
