<?php

namespace App\Exports;

use App\Models\Pembelian;
use Maatwebsite\Excel\Concerns\FromCollection;
use Maatwebsite\Excel\Concerns\WithHeadings;
use Maatwebsite\Excel\Concerns\WithMapping;
use Maatwebsite\Excel\Concerns\ShouldAutoSize;
use Maatwebsite\Excel\Concerns\WithEvents;
use Maatwebsite\Excel\Events\AfterSheet;

class CustomerExport implements FromCollection, WithHeadings, WithMapping, ShouldAutoSize, WithEvents
{
    public function collection()
    {
        return Pembelian::with(['customer', 'details.produk'])->get();
    }

    public function headings(): array
    {
        return [
            'Nama Pelanggan',
            'No HP Pelanggan',
            'Poin Pelanggan',
            'Produk',
            'qty',
            'sub total',
            'Total Harga',
            'Total Bayar',
            'Total Diskon Poin',
            'Total Kembalian',
            'Tanggal Pembelian'
        ];
    }

    public function map($transaction): array
    {
        $rows = [];

        foreach ($transaction->details as $index => $detail) {
            $productName = $detail->produk->nama_produk ?? '-';
            $qty = $detail->quantity;
            $subTotal = 'Rp ' . number_format($detail->sub_total, 2, ',', '.');

            if ($index == 0) {
                $rows[] = [
                    $transaction->customer->nama ?? 'Bukan Member',
                    $transaction->customer->no_hp ?? '-',
                    $transaction->customer->total_point ?? 0,
                    $productName,
                    $qty,
                    $subTotal,
                    'Rp ' . number_format($transaction->total_price, 2, ',', '.'),
                    'Rp ' . number_format($transaction->total_payment, 2, ',', '.'),
                    'Rp ' . number_format($transaction->used_point, 2, ',', '.'),
                    'Rp ' . number_format($transaction->total_return, 2, ',', '.'),
                    $transaction->created_at->format('d-m-Y'),
                ];
            } else {
                $rows[] = [
                    '', '', '', $productName, $qty, $subTotal, '', '', '', '', ''
                ];
            }
        }

        return $rows;
    }

    public function registerEvents(): array
    {
        return [
            AfterSheet::class => function(AfterSheet $event) {
                // Tambahkan judul di atas header
                $event->sheet->insertNewRowBefore(1, 2); // Sisipkan 2 baris kosong di atas

                $event->sheet->mergeCells('A1:K1');
                $event->sheet->setCellValue('A1', 'Laporan Transaksi Pembelian Pelanggan');

                // Tambahkan styling opsional
                $event->sheet->getStyle('A1')->getFont()->setBold(true)->setSize(14);
                $event->sheet->getStyle('A1')->getAlignment()->setHorizontal('center');
            },
        ];
    }
}
