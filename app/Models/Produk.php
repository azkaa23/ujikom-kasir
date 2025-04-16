<?php

            namespace App\Models;
            
            use Illuminate\Database\Eloquent\Factories\HasFactory;
            use Illuminate\Database\Eloquent\Model;
            
            class Produk extends Model
            {
                use HasFactory;
            
                protected $fillable = [
                    'nama_produk',
                    'harga',
                    'stok',
                    'image',
                ];
            
                public function transaction_detail()
                {
                    return $this->hasMany(Transaction_detail::class);
                }
            }