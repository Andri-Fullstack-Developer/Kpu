<?php

use App\Http\Controllers\AdminController;
use App\Http\Controllers\AuthController;
use App\Http\Controllers\ImpportDataContoller;
use App\Http\Controllers\KandidatContoller;
use App\Http\Controllers\KpuContoller;
use App\Http\Controllers\VotesContoller;
use Illuminate\Support\Facades\Auth;
use Illuminate\Support\Facades\Route;

/*
|--------------------------------------------------------------------------
| Web Routes
|--------------------------------------------------------------------------
|
| Here is where you can register web routes for your application. These
| routes are loaded by the RouteServiceProvider and all of them will
| be assigned to the "web" middleware group. Make something great!
|
 */
// Login KPU
Route::get('/', [KpuContoller::class, 'loginKpu'])->name('login');
Route::post('/login', [KpuContoller::class, 'loginKpuCrediensial'])->name('login.submit');

Route::get('kpu', function () {
    if (Auth::check()) {
        return redirect()->route('kpu');
    }
});

Route::middleware('auth')->group(function () {
    Route::get('kpu', [KpuContoller::class, 'kpu'])->name('kpu');
    Route::post('logout.kpu', [KpuContoller::class, 'logoutKpu'])->name('logout.kpu');
    Route::post('vote', [VotesContoller::class, 'store'])->name('vote.store');
    Route::get('/votes', [VotesContoller::class, 'showVotedCandidate'])->name('votes.show');
    Route::post('/cast-vote', [VotesContoller::class, 'castVote'])->name('vote.cast');
});
Route::get('thanks', [KpuContoller::class, 'thanksKpu'])->name('thanksKpu');

// Login Administrator
Route::get('/admin', [AdminController::class, 'loginAdmin'])->name('admin');
Route::post('admin-login', [AdminController::class, 'adminCrediensial'])->name('admin.submit');

Route::get('dashboard', function () {
    if (Auth::check()) {
        return redirect()->route('dashboard');
    }
});

Route::middleware('auth.custom:admin')->group(function () {
    Route::get('/dashboard', [AuthController::class, 'dashboard'])->name('dashboard');

    // Tambah Admin
    Route::get('tambah-admin', [AdminController::class, 'admin'])->name('tambah-admin');
    Route::post('tambah-admin/create', [AdminController::class, 'createAdmin'])->name('createAdmin');
    Route::put('tambah-admin/update/{id}', [AdminController::class, 'update_admin'])->name('update.admin');
    Route::delete('tambah-admin/delete/{id}', [AdminController::class, 'delete_admin'])->name('delete.admin');

    // Tambah User
    Route::get('tambah-user', [AuthController::class, 'tambahUser'])->name('tambah-user');
    Route::post('tambah-user/create', [AuthController::class, 'createUser'])->name('tambah');
    // Lits User
    Route::get('list-users', [AuthController::class, 'listUsers'])->name('listUsers');
    Route::put('update/user/{id}', [AuthController::class, 'updateUser'])->name('editUser.create');
    Route::delete('users/delete', [AuthController::class, 'destroy'])->name('delete');

    //
    Route::post('/logout.admin', [AdminController::class, 'logoutAdmin'])->name('logout.admin');
    // Kandidat
    Route::get('kandidat', [KandidatContoller::class, 'kandidat'])->name('kandidat');
    Route::post('kandidat/create', [KandidatContoller::class, 'createKandidat'])->name('kandidat.create');
    Route::delete('kandidat/delete/{id}', [KandidatContoller::class, 'deleteKandidat'])->name('kandidat.delete');
    Route::get('kandidat/{id}', [KandidatContoller::class, 'model'])->name('kandidat.show');
    Route::put('kandidat/update/{id}', [KandidatContoller::class, 'updateKandidat'])->name('edit.kandidat');
    // Score
    Route::get('score', [KandidatContoller::class, 'scoreAkhir'])->name('scoreAkhir');
    // routes/web.php
    Route::get('/score/{id}', [KandidatContoller::class, 'show'])->name('score.show');

    Route::post('import-data', [ImpportDataContoller::class, 'importData'])->name('importData');

    // Votes
    Route::get('votes', [AuthController::class, 'votesView'])->name('votes');
    Route::delete('votes/delete', [VotesContoller::class, 'deleteVotes'])->name('delete.votes');
});

//
// Route::get('tambah-admin', [AdminController::class, 'admin'])->name('tambah-admin');
// Route::post('tambah-admin/create', [AdminController::class, 'createAdmin' ])->name('createAdmin');

// Route::get('/', function () {
//     return view('welcome');
// });

// Route::post('/login', [AuthController::class, 'loginUser'])->name('login');

// KPU
// Route::get('kpu', [AuthController::class, 'kpu']);

// Route::get('dashboard', [AuthController::class, 'dashboard'])->name('dashboard');

// Login
// Route::get('/kpu', function () {
//     if (Auth::check()) {
//         return redirect()->route('kpu');
//     }
//     return view('login-siswa');
// })->name('login');
//
// Route::middleware('auth')->group(function () {
//     Route::get('kpu', [AuthController::class, 'kpu'])->name('kpu');
// Route::get('dashboard', [AuthController::class, 'dashboard'])->name('dashboard');
// Route::get('tambah-user', [AuthController::class, 'tambahUser'])->name('tambah-user');
// Route::post('tambah-user/create', [AuthController::class, 'createUser'])->name('tambah');
// Route::get('list-users', [AuthController::class, 'listUsers'])->name('listUsers');
//     Route::post('/logout', [AuthController::class, 'logout'])->name('logout');
// });

// Login Admin