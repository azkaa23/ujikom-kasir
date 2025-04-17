@extends('layouts.master')

@section('title', 'Home Page')

@section('content')
<div class="page-wrapper">
    <!-- ============================================================== -->
    <!-- Bread crumb and right sidebar toggle -->
    <!-- ============================================================== -->
    <div class="page-breadcrumb">
        <div class="row align-items-center">
            <div class="col-6">
                <nav aria-label="breadcrumb">
                    <ol class="breadcrumb mb-0 d-flex align-items-center">
                      <li class="breadcrumb-item"><a href="index.html" class="link"><i class="mdi mdi-home-outline fs-4"></i></a></li>
                      <li class="breadcrumb-item active" aria-current="page">Produk<li>
                    </ol>
                  </nav>
                <h1 class="mb-0 fw-bold">Produk</h1> 
            </div>
        </div>
    </div>
    @if (Session::get('success'))
            <div class="alert alert-success">{{ Session::get('success' )}}</div>
        @endif
    <!-- ============================================================== -->
    <!-- End Bread crumb and right sidebar toggle -->
    <!-- ============================================================== -->
    <!-- ============================================================== -->
    <!-- Container fluid  -->
    <!-- ============================================================== -->
    <div class="container-fluid">
        <!-- ============================================================== -->
        <!-- Start Page Content -->
        <!-- ============================================================== -->
        <div class="row">
            <div class="col-12">
                <div class="card">
                    <div class="card-body">
                        @if (Auth::check())
                            @if (Auth::user()->role == 'admin')
                            <div class="col-6">
                                <a href="{{ route('produk.produk.export') }}"  class="btn btn-info">
                                     Export Produk (.xlsx)
                                </a>
                            </div>
                    <div class="d-grid gap-2 d-md-flex justify-content-md-end">
                        <a class="btn btn-primary" href="{{ route('produk.create')}}">Tambah Produk</a>
                      </div>
                      @endif
                        <div class="table-responsive">
                            <table class="table">
                                <thead>
                                    <tr>
                                        <th scope="col">#</th>
                                        <th scope="col"></th>
                                        <th scope="col">Nama</th>
                                        <th scope="col">Harga</th>
                                        <th scope="col">Stok</th>
                                        @if (Auth::user()->role == 'admin')
                                        <th scope="col"></th>
                                        @endif
                                    </tr>
                                </thead>
                                <tbody>
                                    @foreach ($data as $item)
                                        <tr>
                                            <td>{{ $loop->iteration }}</td>
                                            <td><img src="{{ asset('storage/' . $item->image) }}" alt="Produk" width="80"></td>
                                            <td>{{ $item->nama_produk }}</td>
                                            <td>Rp {{ number_format($item->harga, 0, ',', '.') }}</td>
                                            <td>{{ $item->stok }}</td>
                                            @if (Auth::user()->role == 'admin')
                                            <td class="d-flex justify-content-center">
                                                <a href="{{ route('produk.edit', $item->id) }}" class="btn btn-warning mb-3">Edit</a>
                                                <button class="btn btn-info mb-3 ms-1" data-bs-toggle="modal" data-bs-target="#updateStockModal{{ $item->id }}">Update Stok</button>
                                                <form action="{{ route('produk.destroy', $item->id) }}" method="POST" >
                                                    @csrf
                                                    @method('DELETE')
                                                    <button type="submit" class="btn btn-danger mb-3 ms-1" onclick="return confirm('Hapus produk ini?')">Hapus</button>
                                                </form>
                                            </td>
                                            @endif
                                        </tr>  

                                        <div class="modal fade" id="updateStockModal{{ $item->id }}" tabindex="-1" aria-labelledby="updateStockModalLabel{{ $item->id }}" aria-hidden="true">
                                            <div class="modal-dialog">
                                                <div class="modal-content">
                                                    <div class="modal-header">
                                                        <h5 class="modal-title" id="updateStockModalLabel{{ $item->id }}">Update Stok Produk</h5>
                                                        <button type="button" class="btn-close" data-bs-dismiss="modal" aria-label="Close"></button>
                                                    </div>
                                                    <div class="modal-body">
                                                        <form action="{{ route('produk.stok', $item->id) }}" method="POST">
                                                            @csrf
                                                            @method('PUT')
                                                            <div class="mb-3">
                                                                <label for="nama_produk" class="form-label">Nama Produk</label>
                                                                <input type="text" class="form-control" id="nama_produk" name="nama_produk" value="{{ $item->nama_produk }}" readonly>
                                                            </div>
                                                            <div class="mb-3">
                                                                <label for="stok" class="form-label">Stok Baru</label>
                                                                <input type="number" class="form-control" id="stok" name="stok" value="{{ $item->stok }}" required>
                                                            </div>
                                                            <button type="submit" class="btn btn-primary">Update Stok</button>
                                                            <button type="button" class="btn btn-secondary" data-bs-dismiss="modal">Batal</button>
                                                        </form>
                                                    </div>
                                                </div>
                                            </div>
                                        </div>
                                    @endforeach
                                </tbody>
                            </table>
                        </div>
                    </div>
                    @endif
                </div>
        </div>
        <!-- ============================================================== -->
        <!-- End PAge Content -->
        <!-- ============================================================== -->
        <!-- ============================================================== -->
        <!-- Right sidebar -->
        <!-- ============================================================== -->
        <!-- .right-sidebar -->
        <!-- ============================================================== -->
        <!-- End Right sidebar -->
        <!-- ============================================================== -->
    </div>
    <!-- ============================================================== -->
    <!-- End Container fluid  -->
    <!-- ============================================================== -->
    <!-- ============================================================== -->
    <!-- footer -->
    <!-- ============================================================== -->
    <!-- ============================================================== -->
    <!-- End footer -->
    <!-- ============================================================== -->
</div>
@endsection