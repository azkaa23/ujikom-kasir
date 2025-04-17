<?php

namespace App\Exports;

use App\Models\Produk;
use Maatwebsite\Excel\Concerns\FromCollection;
use Maatwebsite\Excel\Concerns\WithHeadings;
use Maatwebsite\Excel\Concerns\WithMapping;
use Maatwebsite\Excel\Concerns\ShouldAutoSize;
use Maatwebsite\Excel\Concerns\WithStyles;
use Maatwebsite\Excel\Concerns\WithEvents;
use PhpOffice\PhpSpreadsheet\Worksheet\Worksheet;
use Maatwebsite\Excel\Events\AfterSheet;

class ProdukExport implements FromCollection, WithHeadings, WithMapping, ShouldAutoSize, WithEvents
{
    public function collection()
    {
        return Produk::all();
    }

    public function headings(): array
    {
        return [
            'No',
            'Nama Produk',
            'Harga',
            'Stok',
         
        ];
    }

    public function map($produk): array
    {
        return [
            $produk->id,
            $produk->nama_produk,
            'Rp ' . number_format($produk->harga, 0, ',', '.'),
            $produk->stok,
            
        ];
    }



    public function registerEvents(): array
    {
        return [
            AfterSheet::class => function(AfterSheet $event) {
                // Sisipkan 2 baris di atas
                $event->sheet->insertNewRowBefore(1, 2);

                // Gabungkan cell baris pertama
                $event->sheet->mergeCells('A1:E1');
                $event->sheet->setCellValue('A1', 'Data Produk Toko');

                // Style untuk judul
                $event->sheet->getStyle('A1')->getFont()->setBold(true)->setSize(14);
                $event->sheet->getStyle('A1')->getAlignment()->setHorizontal('center');
            },
        ];
    }
}
