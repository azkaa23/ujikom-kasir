<?php

            namespace App\Models;
            
            use Illuminate\Database\Eloquent\Factories\HasFactory;
            use Illuminate\Database\Eloquent\Model;
            
            class Transaction_detail extends Model
            {
                use HasFactory;
            
                protected $fillable = [
                    'transaction_id',
                    'produk_id',
                    'quantity',
                    'sub_total',
                ];
            
                public function produk()
                {
                    return $this->belongsTo(Produk::class, 'produk_id');
                }
            
                public function pembelian()
                {
                    return $this->belongsTo(Pembelian::class, 'transaction_id');
                }
            }