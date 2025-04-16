<?php

namespace App\Exports;

use App\Models\TransactionDetail;
use App\Models\Customer;
use App\Models\Pembelian;
use App\Models\Produk;
use Maatwebsite\Excel\Concerns\FromCollection;
use Maatwebsite\Excel\Concerns\WithHeadings;
use Maatwebsite\Excel\Concerns\WithMapping;
use Maatwebsite\Excel\Concerns\ShouldAutoSize;

class CustomerExport implements FromCollection, WithHeadings, WithMapping, ShouldAutoSize
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
            $productName = $detail->produk->nama_produk . ' (' . $detail->quantity . ' : Rp ' . number_format($detail->sub_total, 2, ',', '.') . ')';

            if ($index == 0) {
                $rows[] = [
                    $transaction->customer->nama ?? 'Bukan Member', // Nama Pelanggan
                    $transaction->customer->no_hp ?? '-',          // No HP Pelanggan
                    $transaction->customer->total_point ?? 0,      // Poin Pelanggan
                    $productName,                                  // Produk
                    'Rp ' . number_format($transaction->total_price, 2, ',', '.'), // Total Harga
                    'Rp ' . number_format($transaction->total_payment, 2, ',', '.'), // Total Bayar
                    'Rp ' . number_format($transaction->used_point, 2, ',', '.'),  // Total Diskon Poin
                    'Rp ' . number_format($transaction->total_return, 2, ',', '.'), // Total Kembalian
                    $transaction->created_at->format('d-m-Y'),     // Tanggal Pembelian
                ];
            } else {
                // For other products, just add the product information with empty customer data
                $rows[] = [
                    '', // Empty Nama Pelanggan for subsequent rows
                    '', // Empty No HP Pelanggan for subsequent rows
                    '', // Empty Poin Pelanggan for subsequent rows
                    $productName,                                  // Produk
                    '', // Empty Total Harga for subsequent rows
                    '', // Empty Total Bayar for subsequent rows
                    '', // Empty Total Diskon Poin for subsequent rows
                    '', // Empty Total Kembalian for subsequent rows
                    '', // Empty Tanggal Pembelian for subsequent rows
                ];
            }
        }

        return $rows;
    }
}
