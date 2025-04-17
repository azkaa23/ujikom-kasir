<?php

namespace App\Exports;

use App\Models\User;
use Maatwebsite\Excel\Concerns\FromCollection;
use Maatwebsite\Excel\Concerns\WithHeadings;
use Maatwebsite\Excel\Concerns\WithMapping;
use Maatwebsite\Excel\Concerns\WithEvents;
use Maatwebsite\Excel\Events\AfterSheet;

class UserExport implements FromCollection, WithHeadings, WithMapping, WithEvents
{
    public function collection()
    {
        return User::all();
    }

    public function headings(): array
    {
        return [
            'No',
            'Email',
            'Nama User',
            'Role'
        ];
    }

    public function map($user): array
    {
        return [
            $user->id ?? '-',
            $user->email ?? '-',
            $user->nama ?? '-',
            $user->role ?? '-',
        ];
    }

    public function registerEvents(): array
    {
        return [
            AfterSheet::class => function(AfterSheet $event) {
                // Sisipkan 2 baris di atas
                $event->sheet->insertNewRowBefore(1, 2);

                // Gabungkan cell dan tambahkan judul
                $event->sheet->mergeCells('A1:D1');
                $event->sheet->setCellValue('A1', 'Daftar Pengguna Aplikasi');

                // Styling
                $event->sheet->getStyle('A1')->getFont()->setBold(true)->setSize(14);
                $event->sheet->getStyle('A1')->getAlignment()->setHorizontal('center');
            },
        ];
    }
}
