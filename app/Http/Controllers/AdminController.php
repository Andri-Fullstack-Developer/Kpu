<?php

namespace App\Http\Controllers;

use App\Models\AdminUserModel;
use GuzzleHttp\Promise\Create;
use Illuminate\Http\Request;
use Illuminate\Support\Facades\Auth;
use Illuminate\Support\Facades\Log;

class AdminController extends Controller
{
    public function loginAdmin()
    {
        if (Auth::guard('admin')->check()) {
            return redirect()->route('dashboard');
        }
        return view('login-admin');
    }

    public function admin()
    {
        $admin = Auth::guard('admin')->user();

        $list_admin = AdminUserModel::all();

        return view('dashboard.tambah-admin', compact('list_admin'));
    }

    public function adminCrediensial(Request $req)
    {
        $credentials = $req->validate([
            'username_admin' => 'required|string',
            'password_admin' => 'required|string',
        ]);

        Log::info('Attempting login with credentials: ', $credentials);

        try {
            if (Auth::guard('admin')->attempt([
                'username_admin' => $credentials['username_admin'],
                'password' => $credentials['password_admin'], // Pastikan ini benar
            ])) {
                Log::info('Admin login successful');
                return redirect()->route('dashboard')->with('success', 'Admin login successful');
            } else {
                Log::error('Admin login failed', ['credentials' => $credentials]);
                return back()->withErrors([
                    'username_admin' => 'The provided credentials do not match our records.',
                ])->onlyInput('username_admin');
            }
        } catch (\Illuminate\Database\QueryException $e) {
            Log::error('MySQL error: ' . $e->getMessage());
            return back()->withErrors([
                'mysql_error' => 'Error MySQL: ' . $e->getMessage(),
            ])->onlyInput('username_admin');
        }
    }

    public function createAdmin(Request $req)
    {
        $validated = $req->validate([
            'name-admin' => 'required|string|max:255',
            'username-admin' => 'required|string|max:255',
            'password-admin' => 'required|string|max:255',
        ]);

        $roles = 'Administrator';

        $admin = AdminUserModel::Create([
            'name_admin' => $validated['name-admin'],
            'username_admin' => $validated['username-admin'],
            'password_admin' => bcrypt($validated['password-admin']),
            'role_admin' => $roles,
        ]);

        return redirect()->route('tambah-admin')->with('success', 'User created successfully');
    }

    public function update_admin(Request $req, $id)
    {
        $req->validate([
            'name-admin' => 'required|string|max:255',
            'username-admin' => 'required|string|max:255',
            'password-admin' => 'nullable|string|min:8', // Password tidak wajib diisi
        ]);

        $user = AdminUserModel::findOrFail($id);

        $user->name_admin = $req->input('name-admin');
        $user->username_admin = $req->input('username-admin');

        // Cek apakah password diisi
        if ($req->filled('password-admin')) {
            // Jika password diisi, enkripsi dan simpan
            $user->password = bcrypt($req->input('password-admin'));
        }

        // Simpan perubahan
        $user->save();

        return redirect()->route('tambah-admin')->with('success', 'User berhasil diupdate.');

    }

    public function delete_admin($id)
    {
        $user_admin = AdminUserModel::findOrFail($id);

        $user_admin->delete();

        return redirect()->route('tambah-admin')->with('success', 'Pengguna berhasil dihapus.');
    }
    public function logoutAdmin(Request $req)
    {
        Auth::guard('admin')->logout(); // Pastikan logout dilakukan pada guard 'admin'
        $req->session()->invalidate();
        $req->session()->regenerateToken();

        return redirect()->route('admin');
    }

}