@extends('layouts.master')

@section('title', 'Penjualan - Kasir')
@section('content')
<style>
    .container {
        text-align: center;
        margin-top: 50px;
    }
    .button {
        font-size: 20px;
        cursor: pointer;
        margin: 0 10px;
    }
    .sub-total {
        font-size: 24px;
        font-weight: bold;
    }
    .quantity {
        font-size: 20px;
        margin: 0 10px;
    }
    .floating-button {
            position: fixed;
            bottom: 10px;
            left: 50%;
            transform: translateX(-50%);
            z-index: 1000;
    }
</style>
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
                        <section>
                            <div class="text-center container">
                                <div class="row" id="product-list">
                                    @foreach ($data as $item)
                                    <div class="col-lg-4 col-md-6">
                                        <div class="card">
                                            <div class="bg-image hover-zoom ripple ripple-surface ripple-surface-light">
                                                <img src="{{ asset('storage/' . $item->image) }}" class="w-50" />
                                            </div>
                                            <div class="card-body">
                                                <h5 class="card-title mb-3" id="name_{{$item->id}}">{{$item->nama_produk}}</h5>
                                                <p>Stok <span id="stock_{{$item->id}}">{{$item->stok}}</span></p>
                                                <h6 class="mb-3">Rp. {{ number_format($item->harga, 0, ',', '.') }}</h6>
                                                <p id="price_{{$item->id}}" hidden>{{$item->harga}}</p>
                                                <center>
                                                    <table>
                                                        <tr>
                                                            <td class="minus" id="minus_{{$item->id}}" style="padding: 0px 10px; cursor: pointer;"><b>-</b></td>
                                                            <td class="num" id="quantity_{{$item->id}}" style="padding: 0px 10px;">0</td>
                                                            <td class="plus" id="plus_{{$item->id}}" style="padding: 0px 10px; cursor: pointer;"><b>+</b></td>
                                                        </tr>
                                                    </table>
                                                </center>
                                                <br>
                                                <p id="price_{{$item->id}}" hidden>{{$item->harga}}</p>
                                                <p>Sub Total <b><span id="total_{{$item->id}}">Rp. 0</span></b></p>
                                            </div>
                                        </div>
                                    </div>
                                    @endforeach
                                </div>
                                {{-- <button id="next-button" class="btn btn-primary">Selanjutnya</button> --}}
                            </div>
                        </section>
                    </div>
                    <div class="row fixed-bottom d-flex justify-content-end align-content-center"
                    style="margin-left: 18%; width: 83%; height: 70px; border-top: 3px solid #EEE4B1; background-color: white;">
                    <div class="col text-center" style="margin-right: 50px;">
                        <button id="next-button" class="btn btn-primary">Selanjutnya</button>
                    </div>
                </div>

        </div>
    </div>
</div>

<script src="https://code.jquery.com/jquery-3.6.0.min.js"></script>
<script>
   $(document).ready(function() {
    let selectedProducts = [];

    $(".plus").click(function() {
        let id = $(this).attr("id").split("_")[1];
        let quantityElem = $("#quantity_" + id);
        let totalElem = $("#total_" + id);
        let price = parseInt($("#price_" + id).text());
        let stockElem = $("#stock_" + id);
        let name = $("#name_" + id).text(); 

        let quantity = parseInt(quantityElem.text());
        let stock = parseInt(stockElem.text());

        if (stock > 0) {
            quantity++;
            stock--;
            quantityElem.text(quantity);
            stockElem.text(stock);

            let subtotal = price * quantity;
            totalElem.text("Rp. " + formatRupiah(subtotal));

            let productIndex = selectedProducts.findIndex(p => p.id === id);
            if (productIndex === -1) {
                selectedProducts.push({ id: id, name: name, price: price, quantity: quantity, subtotal: subtotal });
            } else {
                selectedProducts[productIndex].quantity = quantity;
                selectedProducts[productIndex].subtotal = subtotal;
            }
        } else {
            alert("Stok habis!");
        }
    });

    $(".minus").click(function() {
        let id = $(this).attr("id").split("_")[1];
        let quantityElem = $("#quantity_" + id);
        let totalElem = $("#total_" + id);
        let price = parseInt($("#price_" + id).text());
        let stockElem = $("#stock_" + id);

        let quantity = parseInt(quantityElem.text());
        let stock = parseInt(stockElem.text());

        if (quantity > 0) {
            quantity--;
            stock++;
            quantityElem.text(quantity);
            stockElem.text(stock);

            let subtotal = price * quantity;
            totalElem.text("Rp. " + formatRupiah(subtotal));

            let productIndex = selectedProducts.findIndex(p => p.id === id);
            if (productIndex !== -1) {
                if (quantity === 0) {
                    selectedProducts.splice(productIndex, 1);
                } else {
                    selectedProducts[productIndex].quantity = quantity;
                    selectedProducts[productIndex].subtotal = subtotal;
                }
            }
        }
    });

    $("#next-button").click(function() {
        let form = $("<form>", {
            action: "{{ route('pembelian.sale') }}",
            method: "GET",
            style: "display: none;"
        });

        form.append($("<input>", {
            type: "hidden",
            name: "_token",
            value: "{{ csrf_token() }}"
        }));

        selectedProducts.forEach(product => {
            form.append($("<input>", {
                type: "hidden",
                name: "products[]",
                value: product.id + ";" + product.name + ";" + product.price + ";" + product.quantity + ";" + product.subtotal
            }));
        });

        $("body").append(form);
        form.submit();
    });

    function formatRupiah(angka) {
        return angka.toString().replace(/\B(?=(\d{3})+(?!\d))/g, ".");
    }
});

</script>
@endsection