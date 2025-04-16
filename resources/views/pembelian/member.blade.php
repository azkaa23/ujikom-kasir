@extends('layouts.master')

@section('content')
    <div class="page-wrapper">
        <div class="page-breadcrumb">
            <div class="row align-items-center">
                <div class="col-6">
                    <nav aria-label="breadcrumb">
                        <ol class="breadcrumb mb-0 d-flex align-items-center">
                            <li class="breadcrumb-item"><a href="index.html" class="link"><i
                                        class="mdi mdi-home-outline fs-4"></i></a></li>
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
                            <form action="{{ route('pembelian.memberStore')}}" method="POST">
                                @csrf
                                <input type="hidden" name="_token" value="{{ csrf_token() }}" autocomplete="off">
                                <input type="hidden" name="_method">
                                <div class="row">
                                    <div class="col-lg-6 col-md-12">
                                        <div class="table table-bordered">
                                            <table>
                                                <tr class="tabletitle">
                                                    <td class="item">Nama Produk</td>
                                                    <td class="item">QTy</td>
                                                    <td class="item">Harga</td>
                                                    <td class="item">Sub Total</td>
                                                </tr>
                                                @foreach ($transactionDetails as $detail)
                                                    <tr class="service">
                                                        <td class="tableitem">
                                                            <p class="itemtext">{{ $detail->produk->nama_produk }}</p>
                                                        </td>
                                                        <td class="tableitem">
                                                            <p class="itemtext">{{ $detail->quantity }}</p>
                                                        </td>
                                                        <td class="tableitem">
                                                            <p class="itemtext">{{ number_format($detail->produk->harga, 0, ',', '.') }}</p>
                                                        </td>
                                                        <td class="tableitem">
                                                            <p class="itemtext">{{ number_format($detail->sub_total, 0, ',', '.') }}</p>
                                                        </td>
                                                    </tr>
                                                @endforeach
                                                <tr class="tabletitle">
                                                    <td></td>
                                                    <td></td>
                                                    <td><h4>Total Harga</h4></td>
                                                    <td><h4>{{ number_format($transaction->total_price, 0, ',', '.') }}</h4></td>
                                                </tr>
                                                <tr class="tabletitle">
                                                    <td></td>
                                                    <td></td>
                                                    <td><h4>Total Bayar</h4></td>
                                                    <td><h4>{{ number_format($transaction->total_payment, 0, ',', '.') }}</h4></td>
                                                </tr>
                                            </table>
                                        </div>
                                    </div>
                                    <div class="col-lg-6 col-md-12">
                                        <div class="row">
                                            <input type="hidden" name="transaction_id" value="{{ $transaction->id }}">
                                            @if ($customer)
                                                <input type="hidden" name="no_hp" value="{{ $customer->no_hp}}">
                                            @endif
                                            <div class="form-group">
                                                <label for="nama" class="form-label">Nama Member (identitas)</label>
                                                <input type="text" name="nama" id="nama" value="{{ $customer->nama ?? '' }}" class="form-control">
                                            </div>
                                            <div class="form-group">
                                                <label for="poin" class="form-label">Poin</label>
                                                <input type="text" name="total_point" id="poin" value="{{ $customer->total_point }}" disabled class="form-control">
                                            </div>
                                            <div class="form-check ms-4">
                                                <input class="form-check-input" type="checkbox" value="Ya" id="check2" 
                                                       {{ $customer && $transactionCount > 1 ? '' : 'disabled' }} name="check_poin">
                                                <label class="form-check-label" for="check2">
                                                    Gunakan poin
                                                </label>
                                                <small class="text-danger">Poin tidak dapat digunakan pada pembelanjaan pertama.</small>
                                            </div>
                                        </div>
                                        <div class="row text-end">
                                            <div class="col-md-12">
                                                <button class="btn btn-primary" type="submit">Selanjutnya</button>
                                            </div>
                                        </div>
                                    </div>
                                </div>
                            </form>
                        </div>
                    </div>
                </div>
            </div>
        </div>
    </div>
@endsection