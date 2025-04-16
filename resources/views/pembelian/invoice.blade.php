<!DOCTYPE html>
<html lang="id">
<head>
    <meta charset="UTF-8">
    <meta name="viewport" content="width=device-width, initial-scale=1.0">
    <title>Struk Pembelian</title>
    <style>
        body {
            font-family: Arial, sans-serif;
            margin: 20px;
        }
        .container {
            max-width: 300px;
            margin: auto;
            border: 1px solid #000;
            padding: 10px;
        }
        .header {
            text-align: center;
            margin-bottom: 10px;
        }
        .member-info {
            margin-bottom: 10px;
        }
        .member-info p {
            margin: 5px 0;
        }
        table {
            width: 100%;
            border-collapse: collapse;
            margin-bottom: 10px;
        }
        table, th, td {
            border: 1px solid #000;
        }
        th, td {
            padding: 5px;
            text-align: left;
        }
        .total {
            text-align: right;
            margin-bottom: 10px;
        }
        .footer {
            text-align: center;
            margin-top: 10px;
        }
    </style>
</head>
<body>
    <div class="container">
        <div class="header">
            <h2>KASIR</h2>
        </div>
        <div class="member-info">
            <p>Member Status : {{ $transaction->customer ? 'Member' : 'NON-MEMBER' }}</p>
            <p>No. HP : {{ $transaction->customer->no_hp ?? '-'}}</p>
            <p>Bergabung Sejak : {{ $transaction->customer && $transaction->customer->created_at ? date('d F Y', strtotime($transaction->customer->created_at)) : '-' }}</p>
            <p>Poin Member : {{ $transaction->customer->total_point ?? '-' }}</p>
        </div>
        <table>
            <thead>
                <tr>
                    <th>Nama Produk</th>
                    <th>QTy</th>
                    <th>Harga</th>
                    <th>Sub Total</th>
                </tr>
            </thead>
            <tbody>
                @foreach ($transactionDetails as $detail)
                <tr>
                    <td>{{ $detail->produk->nama_produk }}</td>
                    <td>{{ $detail->quantity }}</td>
                    <td>Rp.{{ number_format($detail->produk->harga, 0, ',', '.') }}</td>
                    <td>Rp.{{ number_format($detail->sub_total, 0, ',', '.') }}</td>
                </tr>
                @endforeach
            </tbody>
        </table>
        <div class="total">
            <p><strong>Total Harga</strong></p>
            <p>Rp.{{ number_format($transaction->total_price, 0, ',', '.') }}</p>
            <p><strong>Harga Setelah Poin</strong></p>
            <p>Rp.{{ number_format($transaction->total_price, 0, ',', '.') }}</p>
            <p><strong>Total Kembalian</strong></p>
            <p>Rp.{{ number_format($hargaBeforePoint, 0, ',', '.') }}</p>
        </div>
        <div class="footer">
            <p id="current-time">
                {{ \Carbon\Carbon::now()->timezone('Asia/Jakarta')->translatedFormat('d F Y H:i:s') }} | {{ $transaction->user->nama ?? 'Kasir' }}
            </p>            
            <p>Terima kasih atas pembelian Anda!</p>
        </div>
    </div>
</body>
</html>
