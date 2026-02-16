<?php

use function Pest\Laravel\artisan;

$testFile = 'vendor/rap2hpoutre/fast-excel/tests/test1.xlsx';

it('fails when file does not exist', function () {
    $this->artisan('inspect', ['file' => 'nonexistent.xlsx'])
        ->expectsOutputToContain('File not found')
        ->assertExitCode(1);
});

it('lists all sheet names with --sheets option', function () use ($testFile) {
    $this->artisan('inspect', ['file' => $testFile, '--sheets' => true])
        ->expectsOutputToContain('Available sheets')
        ->expectsOutputToContain('Feuil1')
        ->assertExitCode(0);
});

it('fails when sheet index is out of range', function () use ($testFile) {
    $this->artisan('inspect', ['file' => $testFile, '--sheet' => '99'])
        ->expectsOutputToContain('out of range')
        ->assertExitCode(1);
});

it('fails when sheet name does not exist', function () use ($testFile) {
    $this->artisan('inspect', ['file' => $testFile, '--sheet' => 'NonExistentSheet'])
        ->expectsOutputToContain('not found')
        ->assertExitCode(1);
});

it('analyzes sheet data when --sheet option is provided', function () use ($testFile) {
    $this->artisan('inspect', ['file' => $testFile, '--sheet' => '1'])
        ->expectsOutputToContain('Sheet `Feuil1`')
        ->expectsOutputToContain('Sheet statistics')
        ->expectsOutputToContain('Rows')
        ->expectsOutputToContain('col1')
        ->expectsOutputToContain('col2')
        ->assertExitCode(0);
});

it('shows sheet by name', function () use ($testFile) {
    $this->artisan('inspect', ['file' => $testFile, '--sheet' => 'Feuil1'])
        ->expectsOutputToContain('Sheet `Feuil1`')
        ->assertExitCode(0);
});

it('shows images analysis with --images option', function () use ($testFile) {
    $this->artisan('inspect', ['file' => $testFile, '--sheet' => '1', '--images' => true])
        ->expectsOutputToContain('Images in sheet')
        ->expectsOutputToContain('Total images found')
        ->assertExitCode(0);
});

it('shows filled percentage for columns', function () use ($testFile) {
    $this->artisan('inspect', ['file' => $testFile, '--sheet' => '1'])
        ->expectsOutputToContain('Filled')
        ->assertExitCode(0);
});

it('shows distinct values count', function () use ($testFile) {
    $this->artisan('inspect', ['file' => $testFile, '--sheet' => '1'])
        ->expectsOutputToContain('Distinct')
        ->assertExitCode(0);
});

it('respects memory option', function () use ($testFile) {
    $this->artisan('inspect', ['file' => $testFile, '--sheets' => true, '--memory' => '512'])
        ->assertExitCode(0);
});

it('warns when no sheet is specified without --sheets', function () use ($testFile) {
    $this->artisan('inspect', ['file' => $testFile])
        ->expectsOutputToContain('No sheet specified')
        ->assertExitCode(1);
});

it('handles date formatting in test file', function () {
    $testFile = 'vendor/rap2hpoutre/fast-excel/tests/test-dates.xlsx';
    
    $this->artisan('inspect', ['file' => $testFile, '--sheet' => '1'])
        ->assertExitCode(0);
});
