<?php

            namespace App\Models;
            
            use Illuminate\Database\Eloquent\Factories\HasFactory;
            use Illuminate\Database\Eloquent\Model;
            
            class Pembelian extends Model
            {
                use HasFactory;
                protected $fillable = [
                    'user_id',
                    'customer_id',
                    'total_price',
                    'total_payment',
                    'total_return',
                    'point',
                    'used_point',
                ];
            
                public function user()
                {
                    return $this->belongsTo(User::class);
                }
            
                public function details()
                {
                    return $this->hasMany(Transaction_detail::class, 'transaction_id');
                }
            
                public function customer()
                {
                    return $this->belongsTo(Customer::class, 'customer_id');
                }
            }