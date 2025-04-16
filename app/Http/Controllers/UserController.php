<?php

            namespace App\Http\Controllers;
            
            use Illuminate\Http\Request;
            use Illuminate\Support\Facades\Auth;
            use Illuminate\Support\Facades\Hash;
            use App\Models\User;
            
            class UserController extends Controller
            {
                /**
                 * Display a listing of the resource.
                 */
                public function index()
                {
                    $user = User::all();
                    return view('user.index', compact('user'));
                }
            
                public function loginAuth(request $request)
                {
                    $request->validate([
                        'email'=> 'required',
                        'password'=> 'required',
                    ]);
            
                    $user = $request->only(['email', 'password']);
                    if (Auth::attempt($user)) {
                        return redirect()->route('home.page');
                    }else{
                        return redirect()->back()->with('failed', 'Proses login gagal, silahkan coba kembali dengan data benar!');
                    }
                }
            
                /**
                 * Show the form for creating a new resource.
                 */
                public function create()
                {
                    return view('user.create');
                }
            
                /**
                 * Store a newly created resource in storage.
                 */
                public function store(Request $request)
                {
                    $request->validate([
                        'nama' => 'required',
                        'email' => 'required|email|unique:users,email',
                        'password' => 'required',
                        'role' => 'required',
                    ]);
                
                    User::create([
                        'email' => $request->email,
                        'nama' => $request->nama,
                        'password' => Hash::make($request->password),
                        'role' => $request->role,
                    ]);
                
                    return redirect()->route('user.index')->with('success', 'Anda berhasil menambahkan!');
                }
                
            
                /**
                 * Display the specified resource.
                 */
                public function show(string $id)
                {
                    //
                }
            
                /**
                 * Show the form for editing the specified resource.
                 */
                public function edit(string $id)
                {
                    $user = User::find($id);
                    return view('user.edit', compact('user'));
                }
            
                /**
                 * Update the specified resource in storage.
                 */
                public function update(Request $request, string $id)
                {
                    $user = User::find($id);
            
                    $request->validate([
                        'nama' => 'required',
                        'email' => 'required|email',
                        'role' => 'required',
                    ]);
                
                    if ($request->password) {
                        $user->update([
                            'email' => $request->email,
                            'nama' => $request->nama,
                            'password' => $request->password,
                            'role' => $request->role,
                        ]);
                    } else {
                        $user->update([
                            'email' => $request->email,
                            'nama' => $request->nama,
                            'role' => $request->role,
                        ]);
                    }
            
                    return redirect()->route('user.index')->with('success', 'Akun Berhasil Di Ubah');
                }
            
                /**
                 * Remove the specified resource from storage.
                 */
                public function destroy(string $id)
                {
                    User::find($id)->delete();
                    return redirect()->back()->with('delete', 'Berhasil menghapus data');
                }
            
                public function logout()
                {
                    Auth::logout();
                    return redirect()->route('login')->with('logout', 'Anda Berhasil Logout');
                }
            }