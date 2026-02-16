<?php

$testFile = 'vendor/rap2hpoutre/fast-excel/tests/test1.xlsx';

it('fails when file does not exist', function () {
    $this->artisan('spreadsheet', ['file' => 'nonexistent.xlsx'])
        ->expectsOutputToContain('File not found')
        ->assertExitCode(1);
});

it('lists all sheet names with --sheets option', function () use ($testFile) {
    $this->artisan('spreadsheet', ['file' => $testFile, '--sheets' => true])
        ->expectsOutputToContain('Available sheets')
        ->expectsOutputToContain('Feuil1')
        ->assertExitCode(0);
});

it('fails when sheet index is out of range', function () use ($testFile) {
    $this->artisan('spreadsheet', ['file' => $testFile, '--sheet' => '99'])
        ->expectsOutputToContain('out of range')
        ->assertExitCode(1);
});

it('fails when sheet name does not exist', function () use ($testFile) {
    $this->artisan('spreadsheet', ['file' => $testFile, '--sheet' => 'NonExistentSheet'])
        ->expectsOutputToContain('not found')
        ->assertExitCode(1);
});

it('analyzes sheet data when --sheet option is provided', function () use ($testFile) {
    $this->artisan('spreadsheet', ['file' => $testFile, '--sheet' => '1'])
        ->expectsOutputToContain('Sheet `Feuil1`')
        ->expectsOutputToContain('Sheet statistics')
        ->expectsOutputToContain('Rows')
        ->expectsOutputToContain('col1')
        ->expectsOutputToContain('col2')
        ->assertExitCode(0);
});

it('shows sheet by name', function () use ($testFile) {
    $this->artisan('spreadsheet', ['file' => $testFile, '--sheet' => 'Feuil1'])
        ->expectsOutputToContain('Sheet `Feuil1`')
        ->assertExitCode(0);
});

it('shows images analysis with --images option', function () use ($testFile) {
    $this->artisan('spreadsheet', ['file' => $testFile, '--sheet' => '1', '--images' => true])
        ->expectsOutputToContain('Images in sheet')
        ->expectsOutputToContain('Total images found')
        ->assertExitCode(0);
});

it('shows filled percentage for columns', function () use ($testFile) {
    $this->artisan('spreadsheet', ['file' => $testFile, '--sheet' => '1'])
        ->expectsOutputToContain('Filled')
        ->assertExitCode(0);
});

it('shows distinct values count', function () use ($testFile) {
    $this->artisan('spreadsheet', ['file' => $testFile, '--sheet' => '1'])
        ->expectsOutputToContain('Distinct')
        ->assertExitCode(0);
});

it('respects memory option', function () use ($testFile) {
    $this->artisan('spreadsheet', ['file' => $testFile, '--sheets' => true, '--memory' => '512'])
        ->assertExitCode(0);
});

it('warns when no sheet is specified without --sheets', function () use ($testFile) {
    $this->artisan('spreadsheet', ['file' => $testFile])
        ->expectsOutputToContain('No sheet specified')
        ->assertExitCode(1);
});

it('handles date formatting in test file', function () {
    $testFile = 'vendor/rap2hpoutre/fast-excel/tests/test-dates.xlsx';

    $this->artisan('spreadsheet', ['file' => $testFile, '--sheet' => '1'])
        ->assertExitCode(0);
});

it('generates HTML report', function () {
    $testFile = 'tests/fixtures/test.ods';
    $outputFile = sys_get_temp_dir().'/test-report-'.uniqid().'.html';

    $this->artisan('spreadsheet', [
        'file' => $testFile,
        '--sheet' => '1',
        '--output' => 'html',
        '--output-file' => $outputFile,
    ])
        ->expectsOutputToContain('HTML report saved')
        ->assertExitCode(0);

    expect(file_exists($outputFile))->toBeTrue();
    $content = file_get_contents($outputFile);
    expect($content)->toContain('<!DOCTYPE html');
    expect($content)->toContain('Spreadsheet Report');
    expect($content)->toContain('TestSheet');

    unlink($outputFile);
});

it('generates PDF report', function () {
    $testFile = 'tests/fixtures/test.ods';
    $outputFile = sys_get_temp_dir().'/test-report-'.uniqid().'.pdf';

    $this->artisan('spreadsheet', [
        'file' => $testFile,
        '--sheet' => '1',
        '--output' => 'pdf',
        '--output-file' => $outputFile,
    ])
        ->expectsOutputToContain('PDF report saved')
        ->assertExitCode(0);

    expect(file_exists($outputFile))->toBeTrue();
    expect(filesize($outputFile))->toBeGreaterThan(0);

    // Check PDF magic bytes
    $content = file_get_contents($outputFile);
    expect(str_starts_with($content, '%PDF'))->toBeTrue();

    unlink($outputFile);
});

it('fails when output-file is missing for html output', function () {
    $testFile = 'tests/fixtures/test.ods';

    $this->artisan('spreadsheet', [
        'file' => $testFile,
        '--sheet' => '1',
        '--output' => 'html',
    ])
        ->expectsOutputToContain('--output-file is required')
        ->assertExitCode(1);
});

it('fails when output-file is missing for pdf output', function () {
    $testFile = 'tests/fixtures/test.ods';

    $this->artisan('spreadsheet', [
        'file' => $testFile,
        '--sheet' => '1',
        '--output' => 'pdf',
    ])
        ->expectsOutputToContain('--output-file is required')
        ->assertExitCode(1);
});

it('works with LibreOffice ODS files', function () {
    $testFile = 'tests/fixtures/test.ods';

    $this->artisan('spreadsheet', ['file' => $testFile, '--sheet' => '1'])
        ->expectsOutputToContain('TestSheet')
        ->expectsOutputToContain('Name')
        ->expectsOutputToContain('Value')
        ->assertExitCode(0);
});
