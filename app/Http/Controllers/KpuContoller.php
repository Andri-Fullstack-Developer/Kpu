<?php

namespace App\Http\Controllers;

use App\Models\KandidatModel;
use App\Models\User;
use App\Models\VotesModel;
use Illuminate\Http\Request;
use Illuminate\Support\Facades\Auth;

class KpuContoller extends Controller
{
    public function loginKpu()
    {
        if (Auth::check()) {
            return redirect()->route('kpu');
        }

        return view('login-siswa');
    }
    public function kpu()
    {
        // $kandidat = KandidatModel::all();

        $userId = auth()->id();
        $user = User::find($userId);

        $vote = VotesModel::where('user_id', $userId)->first();

        if ($vote) {
            // Ambil kandidat yang dipilih oleh user
            $kandidat = KandidatModel::where('id', $vote->candidate_id)->get();
        } else {
            // Tampilkan semua kandidat jika user belum memilih
            $kandidat = KandidatModel::all();
        }
        return view('dashboard.kotak-kpu', compact('user', 'kandidat'));
    }

    public function loginKpuCrediensial(Request $req)
    {
        try {
            $credentials = $req->validate([
                'username' => 'required|string',
                'password' => 'required|string',
            ]);

            if (Auth::attempt($credentials)) {
                // Cek status pengguna setelah login berhasil
                $user = Auth::user();
                if ($user->status == 'tidak aktif') {
                    // Jika status pengguna "Tidak aktif", logout dan tampilkan pesan kesalahan
                    Auth::logout();
                    return redirect()->back()->with('alert', 'Akun Anda tidak aktif. Silakan hubungi admin.');
                }

                return redirect()->route('kpu')->with('success', 'Login successful');
            }

            // Cari user dengan username yang sama
            $user = User::where('username', $req->input('username'))->first();

            if (!$user) {
                // Jika user tidak ditemukan, tampilkan error username tidak ada
                return redirect()->back()->with('alert', 'Username tidak ada');
            }

            // Jika user ditemukan, maka password salah
            return redirect()->back()->with('alert', 'Password salah');
        } catch (\Illuminate\Database\QueryException $e) {
            return redirect()->back()->with('alert', 'Error MySQL: ' . $e->getMessage());
        }
    }

    public function logoutKpu(Request $req)
    {
        Auth::logout();
//         $req->session()->invalidate();
//
//         $req->session()->regenerateToken();

        return redirect('thanksKpu');
    }

    public function thanksKpu()
    {
        return view('dashboard.thanks');
    }
}