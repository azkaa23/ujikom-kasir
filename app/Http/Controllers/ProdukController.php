<?php

namespace App\Http\Controllers;

use App\Exports\ProdukExport;
use Illuminate\Support\Facades\Storage;
use Illuminate\Http\Request;
use App\Models\Produk;
use Illuminate\Support\Facades\File;
use Maatwebsite\Excel\Facades\Excel;

class ProdukController extends Controller
{
    /**
     * Display a listing of the resource.
     */
    public function index()
    {
        $data = Produk::all();
        return view('produk.index', ['data' => $data]);
    }

    /**
     * Show the form for creating a new resource.
     */
    public function create()
    {
        return view('produk.create');
    }

    /**
     * Store a newly created resource in storage.
     */
    public function store(Request $request)
    {
        // dd($request->all());
        $request->validate([
            'nama_produk' => 'required',
            'harga' => 'required',
            'stok' => 'required',
            'image' => 'required', 
        ]);
    
        $harga = str_replace(['Rp', '.'], '', $request->harga);

        if ($request->hasFile('image')) {
            $extension = $request->file('image')->getClientOriginalExtension();
            $newName = $request->nama_produk . '-' . now()->timestamp . '.' . $extension;
            $path = $request->file('image')->storeAs('uploads', $newName, 'public');
    
            Produk::create([
                'nama_produk' => $request->nama_produk,
                'harga' => $harga, 
                'stok' => $request->stok,
                'image' => $path,
            ]);
    
            return redirect()->route('produk.index')->with('success', 'Anda berhasil menambahkan!');
        }
    
        return redirect()->route('produk.index')->with('error', 'Gagal mengupload gambar.');
    }

    /**
     * Show the form for editing the specified resource.
     */
    public function edit(string $id)
    {
        $produk = Produk::find($id);
        return view('produk.edit', compact('produk'));
    }

    /**
     * Update the specified resource in storage.
     */
    public function update(Request $request, string $id)
    {
        $produk = Produk::find($id);
    
        $request->validate([
            'nama_produk' => 'required',
            'harga' => 'required',
            'stok' => 'required',
        ]);
    
        if ($request->hasFile('image')) {
            $extension = $request->file('image')->getClientOriginalExtension();
            $newName = $request->nama_produk . '-' . now()->timestamp . '.' . $extension;
            $path = $request->file('image')->storeAs('uploads', $newName, 'public');
    
            // Hapus gambar lama jika ada
            if ($produk->image) {
                Storage::disk('public')->delete($produk->image);
            }
    
            $produk->image = 'uploads/' . $newName; // Simpan path baru ke produk
        }
    
        $produk->update([
            'nama_produk' => $request->nama_produk,
            'harga' => (int) str_replace(['.', 'Rp', ' '], '', $request->harga),
            'stok' => $request->stok,
            'image' => $produk->image ?? $produk->getOriginal('image'),
        ]);
        
    
        return redirect()->route('produk.index')->with('success', 'Produk berhasil diperbarui');
    }

    /**
     * Remove the specified resource from storage.
     */
    public function destroy(string $id)
    {
        $produk = Produk::find($id);

        if (!$produk) {
            return redirect()->back()->with('error', 'Produk tidak ditemukan!');
        }

        if ($produk->image) {
            $imagePath = public_path('app/public/uploads/' . $produk->image); 
            
            if (File::exists($imagePath)) {
                File::delete($imagePath); 
            }
        }

        $produk->delete();

        return redirect()->back()->with('delete', 'Berhasil menghapus data dan gambar!');
    }

    public function updateStock(Request $request, $id)
    {
        $request->validate([
            'stok' => 'required|integer|min:0', 
        ]);

        $produk = Produk::findOrFail($id);

        $produk->stok = $request->stok;
        $produk->save();

        return redirect()->route('produk.index')->with('success', 'Stok berhasil diperbarui!');
    }

    public function export() 
    {
        return Excel::download(new ProdukExport, 'data-produk.xlsx');
    }
}
