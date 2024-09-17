<?php

namespace App\Http\Controllers;

use App\Models\KandidatModel;
use App\Models\User;
use App\Models\VotesModel;
use Illuminate\Http\Request;
use Illuminate\Support\Facades\Auth;
use Illuminate\Support\Facades\DB;

class AuthController extends Controller
{

    // public function __construct()
    // {
    //     $this->middleware('auth')->except(['login']);
    // }

    public function loginSiswa()
    {
        if (Auth::check()) {
            return redirect()->route('kpu');
        }

        return view('login-siswa');
    }
    public function login(Request $request)
    {

    }

    // KPU
    public function kpu()
    {
        return view('dashboard.kotak-kpu');
    }

    public function dashboard()
    {
        $admin = Auth::guard('admin')->user();

        $kandidatData = KandidatModel::withCount('votes')
            ->get()
            ->map(function ($kandidat) {
                return [
                    'name' => $kandidat->name_calon,
                    'votes' => $kandidat->votes_count,
                ];
            });

        $pemilihPerKelas = User::select('kelas', DB::raw('count(id) as jumlah_pemilih'))
            ->groupBy('kelas')
            ->get()
            ->map(function ($item) {
                return [
                    'kelas' => $item->kelas,
                    'jumlah_pemilih' => $item->jumlah_pemilih,
                ];
            });

        // $combinedData = DB::table('votes')
        //     ->select(
        //         DB::raw("DATE_FORMAT(votes.vote_time, '%Y-%m-%d %H:%i:%s') as date"),
        //         DB::raw('count(votes.id) as total_votes'),
        //         DB::raw('SUM(CASE WHEN users.status = "aktif" THEN 1 ELSE 0 END) as active_users'),
        //         DB::raw('SUM(CASE WHEN users.status = "tidak aktif" THEN 1 ELSE 0 END) as inactive_users')
        //     )
        //     ->join('users', 'users.id', '=', 'votes.user_id')
        //     ->groupBy(DB::raw("DATE_FORMAT(votes.vote_time, '%Y-%m-%d %H:%i:%s')"))
        //     ->orderBy(DB::raw("DATE_FORMAT(votes.vote_time, '%Y-%m-%d %H:%i:%s')"))
        //     ->get()
        //     ->map(function ($item) {
        //         return [
        //             'date' => $item->date,
        //             'total_votes' => $item->total_votes,
        //             'active_users' => $item->active_users,
        //             'inactive_users' => $item->inactive_users,
        //         ];
        //     });
        $votesData = VotesModel::select(DB::raw('DATE_FORMAT(vote_time, "%H:%i") as minute'), DB::raw('count(*) as total_votes'))
            ->groupBy('minute')
            ->orderBy('minute', 'ASC')
            ->get();

        $groupedBarData = DB::table('users')
            ->select('users.kelas', 'kandidats.name_calon', DB::raw('count(votes.id) as votes'))
            ->join('votes', 'users.id', '=', 'votes.user_id')
            ->join('kandidats', 'votes.candidate_id', '=', 'kandidats.id')
            ->groupBy('users.kelas', 'kandidats.name_calon')
            ->orderBy('users.kelas')
            ->get()
            ->groupBy('kelas')
            ->map(function ($group) {
                $candidates = [];
                foreach ($group as $item) {
                    $candidates[] = [
                        'name_calon' => $item->name_calon,
                        'votes' => $item->votes,
                    ];
                }
                return [
                    'kelas' => $group->first()->kelas,
                    'candidates' => $candidates,
                ];
            })
            ->values();

        $recentVotes = DB::table('votes')
            ->join('users', 'votes.user_id', '=', 'users.id')
            ->join('kandidats', 'votes.candidate_id', '=', 'kandidats.id')
            ->select('users.name as user_name', 'kandidats.name_calon', 'votes.vote_time')
            ->orderBy('votes.vote_time', 'desc')
            ->take(5)
            ->get();

        return view('dashboard.dashboard', [
            'role' => $admin->roles_admin,
            'kandidatData' => $kandidatData,
            'pemilihPerKelas' => $pemilihPerKelas,
            'votesData' => $votesData,
            'groupedBarData' => $groupedBarData,
            'recentVotes' => $recentVotes,
        ]);
    }

    public function tambahUser()
    {
        return view('dashboard.tambah-user');
    }

    public function createUser(Request $req)
    {

        $validated = $req->validate([
            'name' => 'required|string|max:255',
            'username' => 'required|string|max:255',
            'kelas' => 'required|string|max:255',
            'password' => 'required|string|max:255',
        ]);

        $status = 'aktif';

        $user = User::create([
            'name' => $validated['name'],
            'username' => $validated['username'],
            'kelas' => $validated['kelas'],
            'password' => bcrypt($validated['password']),
            'status' => $status,
        ]);

        return redirect()->route('tambah-user')->with('success', 'User created successfully');
    }
    public function loginUser(Request $request)
    {
        $credentials = $request->validate([
            'username' => 'required|string',
            'password' => 'required|string',
        ]);

        if (Auth::attempt($credentials)) {

            return redirect()->intended('dashboard')->with('success', 'Login successful');
        }

        return back()->withErrors([
            'username' => 'The provided credentials do not match our records.',
        ])->onlyInput('username');
    }

    public function listUsers()
    {
        $users = User::all();
        return view('dashboard.list-tabel', compact('users'));
    }
    public function updateUser(Request $req, $id)
    {
        $user = User::find($id);

        $validate = $req->validate([
            'name' => 'required',
            'kelas' => 'required',
        ]);

        $user->name = $req->input('name');
        $user->kelas = $req->input('kelas');

        // Simpan perubahan ke database
        $user->save();

        // Redirect dengan pesan sukses
        return redirect()->route('listUsers')->with('success', 'User updated successfully!');

    }

    public function logout(Request $request)
    {
        Auth::logout();
        $request->session()->invalidate();

        $request->session()->regenerateToken();

        return redirect()->route('login');
    }

    public function destroy(Request $req)
    {

        $ids = $req->input('selected_users', []);

        // Dump and die to check the retrieved IDs
        // dd($ids);

        // Ensure $ids is an array and not empty
        if (is_array($ids) && !empty($ids)) {
            User::whereIn('id', $ids)->delete();
            return redirect()->back()->with('success', 'Users deleted successfully.');
        }

        return redirect()->back()->with('error', 'No users selected.');
    }

    public function votesView()
    {
        $vote = VotesModel::with(['user', 'kandidat'])->get();
        return view('dashboard.votes', compact('vote'));
    }

}
