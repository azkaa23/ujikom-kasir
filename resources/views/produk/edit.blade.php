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
                      <li class="breadcrumb-item"><a href="" class="link">produk</a></li>
                      <li class="breadcrumb-item active" aria-current="page">Tambah produk<li>
                    </ol>
                  </nav>
                <h1 class="mb-0 fw-bold">Tambah produk</h1> 
            </div>
        </div>
    </div>
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
                      <div class="container mt-5">
                        <form action="{{ route('produk.update', $produk->id )}}" method="POST" enctype="multipart/form-data">
                            @csrf
                            @method('PATCH')
                            <div class="row mb-3">
                                <div class="col-md-6">
                                    <label for="nama_produk" class="form-label">Nama Produk</label>
                                    <input type="text" name="nama_produk" class="form-control" id="nama_produk" value="{{$produk->nama_produk}}">
                                </div>
                                <div class="col-md-6">
                                    <label for="harga" class="form-label">Harga</label>
                                    <input type="text" name="harga" class="form-control" id="harga" 
                                    value="Rp. {{ number_format( $produk->harga, 0, ',', '.') }}">
                                </div>
                            </div>
                            <div class="row mb-3">
                                <div class="col-md-6">
                                    <label for="stok" class="form-label">Stok</label>
                                    <input type="stok" name="stok" class="form-control" id="stok" value="{{ $produk->stok }}">
                                </div>
                                <div class="col-md-6">
                                    <label for="image" class="form-label">Role</label>
                                    <input type="file" class="form-control" name="image" id="image">
                                    @if ($produk->image)
                                        <img src="{{ asset('storage/' . $produk->image) }}" alt="Gambar Produk" width="100">
                                    @endif
                                </div>
                            </div>
                            <button type="submit" class="btn btn-primary">Simpan</button>
                        </form>
                    </div>
                    </div>
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

<script>
    document.addEventListener("DOMContentLoaded", function() {
        let priceInput = document.getElementById("harga");

        priceInput.addEventListener("keyup", function(e) {
            this.value = formatRupiah(this.value, "Rp. ");
        });

        document.querySelector("form").addEventListener("submit", function() {
        priceInput.value = priceInput.value.replace(/[^0-9]/g, ""); // Hilangkan karakter selain angka sebelum submit
        });

        function formatRupiah(angka, prefix) {
            let number_string = angka.replace(/[^,\d]/g, "").toString(),
                split = number_string.split(","),
                sisa = split[0].length % 3,
                rupiah = split[0].substr(0, sisa),
                ribuan = split[0].substr(sisa).match(/\d{3}/gi);

            if (ribuan) {
                let separator = sisa ? "." : "";
                rupiah += separator + ribuan.join(".");
            }

            rupiah = split[1] !== undefined ? rupiah + "," + split[1] : rupiah;
            return prefix === undefined ? rupiah : (rupiah ? prefix + rupiah : "");
        }
    });
</script>

@endsection