<?php

            namespace App\Models;
            
            use Illuminate\Database\Eloquent\Factories\HasFactory;
            use Illuminate\Database\Eloquent\Model;
            
            class Customer extends Model
            {
                use HasFactory;
            
                protected $fillable = [
                    'nama',
                    'no_hp',
                    'total_point',
                ];
            
                public function transactions()
                {
                    return $this->hasMany(Pembelian::class);
                }
            }