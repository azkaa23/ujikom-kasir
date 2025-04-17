@extends('layouts.master')

@section('title', 'Home Page')

@section('content')
<div class="page-wrapper">
    <div class="page-breadcrumb">
        <div class="row align-items-center">
            <div class="col-6">
                <nav aria-label="breadcrumb">
                    <ol class="breadcrumb mb-0 d-flex align-items-center">
                        <li class="breadcrumb-item"><a href="index.html" class="link"><i class="mdi mdi-home-outline fs-4"></i></a></li>
                        <li class="breadcrumb-item active" aria-current="page">Penjualan</li>
                    </ol>
                </nav>
                <h1 class="mb-0 fw-bold">Penjualan</h1>
            </div>
        </div>
    </div>

    <div class="container-fluid">
        <div class="row">
            <div class="col-12">
                <div class="card">
                    <div class="card-body">
                        <!-- Tombol dan Form Filter -->
                        <div class="row justify-content-end mb-3">
                            <div class="col text-start">
                                <div class="row">
                                    <div class="col-6">
                                        <a href="{{ route('pembelian.exportExcel')}}" class="btn btn-info">
                                            Export Penjualan (.xlsx)
                                        </a>
                                    </div>
                                    @if (Auth::user()->role == 'employe')
                                    <div class="col text-end">
                                        <a href="{{ route('pembelian.create')}}" class="btn btn-primary">
                                            Tambah Penjualan
                                        </a>
                                    </div>
                                    @endif
                                </div>
                            </div>
                        </div>

                        <div class="mb-3 d-flex justify-content-between align-items-center">
                            <!-- Form untuk per_page -->
                            <form method="GET" action="{{ route('pembelian.index') }}" class="d-flex align-items-center">
                                <label for="entries" class="form-label me-2 mb-0">Tampilkan</label>
                                <select id="entries" name="per_page" class="form-select me-2" onchange="this.form.submit()">
                                    <option value="10" {{ request('per_page') == 10 ? 'selected' : '' }}>10</option>
                                    <option value="25" {{ request('per_page') == 25 ? 'selected' : '' }}>25</option>
                                    <option value="50" {{ request('per_page') == 50 ? 'selected' : '' }}>50</option>
                                </select>
                                @if (request('search'))
                                    <input type="hidden" name="search" value="{{ request('search') }}">
                                @endif
                            </form>

                            <!-- Form untuk search -->
                            <form method="GET" action="{{ route('pembelian.index') }}" class="d-flex">
                                <input type="text" name="search" class="form-control w-auto d-inline-block me-2" placeholder="Cari nama member..." value="{{ request('search') }}">
                                <input type="hidden" name="per_page" value="{{ request('per_page', 10) }}">
                                <button type="submit" class="btn btn-primary btn-sm">Cari</button>
                                @if(request('search'))
                                    <a href="{{ route('pembelian.index') }}" class="btn btn-secondary btn-sm ms-2">Reset</a>
                                @endif
                            </form>
                        </div>

                        <!-- Tabel Penjualan -->
                        <div class="table-responsive">
                            <table id="salesTable" class="table">
                                <thead>
                                    <tr>
                                        <th scope="col">#</th>
                                        <th scope="col">Nama Pelanggan</th>
                                        <th scope="col">Tanggal Penjualan</th>
                                        <th scope="col">Total Harga</th>
                                        <th scope="col">Dibuat Oleh</th>
                                        <th scope="col"></th>
                                    </tr>
                                </thead>
                                <tbody>
                                    @foreach ($sale as $item)
                                    <tr>
                                        <th>{{ $sale->firstItem() + $loop->index }}</th>
                                        <td>{{ $item->customer ? $item->customer->nama : 'NON-MEMBER' }}</td>
                                        <td>{{ date('d-m-Y', strtotime($item->created_at)) }}</td>
                                        <td>Rp {{ number_format($item->total_price, 0, ',', '.') }}</td>
                                        <td>{{ $item->user ? $item->user->nama : 'Tidak Diketahui' }}</td>
                                        <td>
                                            <div class="d-flex justify-content-around">
                                                <button class="btn btn-warning" data-bs-toggle="modal" data-bs-target="#lihat-{{ $item->id }}">Lihat</button>
                                                <form action="{{ route('pembelian.export.pdf', $item->id)}}" method="post">
                                                    @csrf
                                                    <button type="submit" class="btn btn-info">Unduh Bukti</button>
                                                </form>
                                            </div>
                                        </td>
                                    </tr>
                                    @endforeach
                                </tbody>
                            </table>

                            <!-- Modal Detail Penjualan -->
                            @foreach($sale as $item)
                            <div class="modal fade" id="lihat-{{ $item->id }}" tabindex="-1" aria-labelledby="modalLihat" aria-hidden="true">
                                <div class="modal-dialog modal-dialog-scrollable">
                                    <div class="modal-content">
                                        <div class="modal-header">
                                            <h5 class="modal-title" id="modalLihat">Detail Penjualan</h5>
                                            <button type="button" class="btn-close" data-bs-dismiss="modal" aria-label="Close"></button>
                                        </div>
                                        <div class="modal-body">
                                            <div class="row">
                                                <div class="col-6">
                                                    <small>
                                                        <p>Member Status : {{ $item->customer ? 'Member' : 'NON-MEMBER' }}</p>
                                                        <p>No. HP : {{ $item->customer->no_hp ?? '-' }}</p>
                                                        <p>Poin Member : {{ $item->customer->total_point ?? '-' }}</p>
                                                    </small>
                                                </div>
                                                <div class="col-6">
                                                    <small>
                                                        Bergabung Sejak : {{ $item->customer ? date('d F Y', strtotime($item->customer->created_at)) : '-' }}
                                                    </small>
                                                </div>
                                            </div>
                                            <div class="row mb-3 text-center mt-5">
                                                <div class="col-3"><b>Nama Produk</b></div>
                                                <div class="col-3"><b>Qty</b></div>
                                                <div class="col-3"><b>Harga</b></div>
                                                <div class="col-3"><b>Sub Total</b></div>
                                            </div>
                                            @foreach ($item->details as $detail)
                                            <div class="row mb-3">
                                                <div class="col-3">{{ $detail->produk ? $detail->produk->nama_produk : 'Produk Sudah di hapus' }}</div>
                                                <div class="col-3">{{ $detail->quantity }}</div>
                                                <div class="col-3">Rp.{{ $detail->produk && $detail->produk->harga ? number_format($detail->produk->harga, 0, ',', '.') : '0' }}</div>
                                                <div class="col-3">Rp.{{ number_format($detail->sub_total, 0, ',', '.') }}</div>
                                            </div>
                                            @endforeach
                                            <div class="row text-center mt-3">
                                                <div class="col-9 text-end"><b>Total</b></div>
                                                <div class="col-3"><b>Rp.{{ number_format($item->total_price, 0, ',', '.') }}</b></div>
                                            </div>
                                            <div class="row mt-3">
                                                <center>
                                                    Dibuat pada :  {{ \Carbon\Carbon::now()->timezone('Asia/Jakarta')->translatedFormat('d F Y H:i:s') }}  <br> Oleh : {{ $item->user ? $item->user->nama : 'Tidak Diketahui' }}
                                                   
                                                    
                                                             
                                                </center>
                                            </div>
                                        </div>
                                        <div class="modal-footer">
                                            <button type="button" class="btn btn-secondary" data-bs-dismiss="modal">Tutup</button>
                                        </div>
                                    </div>
                                </div>
                            </div>
                            @endforeach

                            <div class="d-flex justify-content-between align-items-center mt-3">
                                <div>
                                    Menampilkan {{ $sale->firstItem() }} - {{ $sale->lastItem() }} dari total {{ $sale->total() }} data
                                </div>
                            
                                <div>
                                    {{ $sale->appends(request()->except('page'))->links('pagination::bootstrap-5') }}
                                </div>
                            </div>
                        </div> <!-- /table-responsive -->
                    </div> <!-- /card-body -->
                </div> <!-- /card -->
            </div>
        </div>
    </div>
</div>
@endsection