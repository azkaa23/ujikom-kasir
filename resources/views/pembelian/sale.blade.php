@extends('layouts.master')

@section('title', 'Penjualan - Kasir')
@section('content')
<script>
    function toggleMemberInput() {
        var memberStatus = document.getElementById("memberStatus").value;
        var memberInput = document.getElementById("memberInput");
        if (memberStatus === "Member") {
            memberInput.style.display = "block";
        } else {
            memberInput.style.display = "none";
        }
    }
</script>
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
                      <li class="breadcrumb-item active" aria-current="page">Penjualan<li>
                    </ol>
                  </nav>
                <h1 class="mb-0 fw-bold">Penjualan</h1> 
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
        <div class="row">
            <div class="col-12">
                <div class="card">
                    <div class="card-body">
                        @if ($errors->any())
                            @foreach ($errors as $error)
                                {{$error}}
                            @endforeach
                        @endif
                        <form action="{{ route('pembelian.store')}}" method="POST">
                            @csrf
                            <div class="row">
                                <div class="col-lg-6 col-md-12">
                                    <table style="width: 100%;">
                                        <thead>
                                            <h2>Produk yang dipilih</h2>
                                        </thead>
                                        <tbody>
                                            @foreach ($products as $product)
                                            <input type="hidden" name="products[]" value="{{ $product['id'] }};{{ $product['name'] }};{{ $product['price'] }};{{ $product['quantity'] }};{{ $product['subtotal'] }};" hidden>
                                            <tr>
                                                <td>{{ $product['name'] }} <br> <small>Rp. {{ number_format($product['price'], 0, ',', '.') }} X {{ $product['quantity'] }}</small></td>
                                                <td><b>Rp. {{ number_format($product['subtotal'], 0, ',', '.') }}</b></td>
                                            </tr>
                                            @endforeach
                                            <tr>
                                                <td style="padding-top: 20px; font-size: 20px;"><b>Total</b></td>
                                                <td class="tex-end" style="padding-top: 20px; font-size: 20px;"><b>Rp. {{ number_format($total, 0, ',', '.') }}</b></td>
                                            </tr>
                                        </tbody>
                                    </table>
                                    <input type="text" name="total" id="total" value="{{ $total }}" hidden>
                                </div>
                                <div class="col-lg-6 col-md-12">
                                    <div class="row">
                                        <div class="form-group">
                                            <label for="member" class="form-label">Member Status</label>
                                            <small class="text-danger">Dapat juga membuat member</small>
                                            <select name="member" id="member" class="form-select" onchange="memberDetect()">
                                                <option value="Bukan Member">Bukan Member</option>
                                                <option value="Member">Member</option>
                                            </select>
                                        </div>
                                    </div>
                                    <div id="member-wrap" class="d-none">
                                        <div class="row">
                                            <div class="col-md-12">
                                                <div class="form-group">
                                                    <label class="col-md-12">No Telepon <small class="text-danger">(daftar/gunakan member)</small></label>
                                                    <div class="col-md-12">
                                                        <input type="number" name="no_hp" class="form-control form-control-line" onKeyPress="if(this.value.length==12) return false;">
                                                    </div>
                                                </div>
                                            </div>
                                        </div>
                                    </div>
                                    <div class="row">
                                        <div class="form-group">
                                            <label for="total_pay" class="form-label">Total Bayar</label>
                                            <input type="text" name="total_pay" id="total_pay" class="form-control" oninput="formatRupiah(this); checkTotalPay()">
                                            <small id="error-message" class="text-danger d-none">Jumlah bayar kurang.</small>
                                        </div>
                                    </div>
                                    <div class="row text-end">
                                        <div class="col-md-12">
                                            <button class="btn btn-primary" id="submit-button" type="submit">Pesan</button>
                                        </div>
                                    </div>
                                </div>
                            </div>
                        </form>
                    </div>
        </div>
    </div>
</div>

<script>
    function formatRupiah(input) {
    let value = input.value.replace(/\D/g, "");
    input.value = value ? 'Rp. ' + new Intl.NumberFormat('id-ID').format(value) : '';
    }

    function checkTotalPay() {
    let total = parseInt(document.getElementById('total').value.replace(/\D/g, "")) || 0;
    let totalPay = parseInt(document.getElementById('total_pay').value.replace(/\D/g, "")) || 0;
    let errorMessage = document.getElementById('error-message');
    let submitButton = document.getElementById('submit-button');

    if (totalPay < total) {
        errorMessage.classList.remove('d-none');
        submitButton.disabled = true; // Disable submit button
    } else {
        errorMessage.classList.add('d-none');
        submitButton.disabled = false; // Enable submit button
    }
}

// Add this event listener to form submission to prevent form if not enough money
document.querySelector('form').addEventListener('submit', function(event) {
    let total = parseInt(document.getElementById('total').value.replace(/\D/g, "")) || 0;
    let totalPay = parseInt(document.getElementById('total_pay').value.replace(/\D/g, "")) || 0;

    if (totalPay < total) {
        event.preventDefault(); // Prevent form submission
        alert('Jumlah bayar kurang!'); // Show an alert if the payment is less than total
    }
});

    function memberDetect() {
        let memberWrap = document.getElementById('member-wrap');
        let memberSelect = document.getElementById('member');

        if (memberSelect.value === 'Member') {
            memberWrap.classList.remove('d-none');
        } else {
            memberWrap.classList.add('d-none');
        }
    }


   
</script>
@endsection