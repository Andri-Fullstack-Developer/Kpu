<?php

namespace App\Http\Controllers;

use App\Models\KandidatModel;
use App\Models\VotesModel;
use Illuminate\Http\RedirectResponse;
use Illuminate\Http\Request;

class KandidatContoller extends Controller
{
    public function kandidat()
    {
        $kandidat = KandidatModel::all();

        return view('dashboard.tambah-kandidat', compact('kandidat'));
    }

    public function model($id)
    {
        $kandidat = KandidatModel::find($id);
        if (!$kandidat) {
            return redirect()->back()->with('error', 'Kandidat tidak ditemukan.');
        }

        session()->flash('modal', $kandidat);
        return redirect()->route('kandidat');
    }
    public function createKandidat(Request $req)
    {
        $validated = $req->validate([
            'image_calon' => 'required|file|mimes:jpg,jpeg,png',
            'name_calon' => 'required',
            'kelas_calon' => 'required',
            'visi_calon' => 'required',
            'misi_calon' => 'required',
        ]);

        try {
            $image = $req->file('image_calon');
            date_default_timezone_set('Asia/Jakarta');
            $imageName = date('HisdmY') . '.' . $image->getClientOriginalExtension();
            $image->move(public_path('images'), $imageName);

            KandidatModel::create([
                'name_calon' => $req->name_calon,
                'kelas_calon' => $req->kelas_calon,
                'visi_calon' => $req->visi_calon,
                'misi_calon' => $req->misi_calon,
                'image_calon' => $imageName,
            ]);

            return redirect()->route('kandidat')->with('success', 'Kandidat berhasil ditambahkan!');
        } catch (\Exception $e) {
            return redirect()->back()->with('error', $e->getMessage());
        }
    }

    public function updateKandidat(Request $req, $id)
    {

        $kandidat = KandidatModel::find($id);

        $validated = $req->validate([
            'image_calon' => 'nullable|file|mimes:jpg,jpeg,png',
            'name_calon' => 'required',
            'kelas_calon' => 'required',
            'visi_calon' => 'required',
            'misi_calon' => 'required',
        ]);

        // dd($validated);

        try {
            // Handle image upload
            if ($req->hasFile('image_calon')) {
                // Delete old image if exists
                if ($kandidat->image_calon && file_exists(public_path('images/' . $kandidat->image_calon))) {
                    unlink(public_path('images/' . $kandidat->image_calon));
                }

                // Store new image
                $image = $req->file('image_calon');
                date_default_timezone_set('Asia/Jakarta');
                $imageName = date('HisdmY') . '.' . $image->getClientOriginalExtension();
                $image->move(public_path('images'), $imageName);

                // Update image name in database
                $kandidat->image_calon = $imageName;
            }

            // Update other fields
            $kandidat->name_calon = $req->input('name_calon');
            $kandidat->kelas_calon = $req->input('kelas_calon');
            $kandidat->visi_calon = $req->input('visi_calon');
            $kandidat->misi_calon = $req->input('misi_calon');

            // Save changes to database
            $kandidat->save();

            return redirect()->back()->with('success', 'Kandidat updated successfully.');
        } catch (\Exception $e) {
            return redirect()->back()->with('error', $e->getMessage());
        }

    }
    public function deleteKandidat($id): RedirectResponse
    {
        try {
            $kandidat = KandidatModel::findOrFail($id);

            $votesCount = VotesModel::where('candidate_id', $kandidat->kodandidat)->count();

            if ($votesCount > 0) {
                return redirect()->back()->with('error', 'Gagal menghapus! Silakan hapus data terkait di tabel votes terlebih dahulu.');
            }

            // Hapus kandidat dari database
            $kandidat->delete();

            // Hapus gambar dari server
            $imagePath = public_path('images/' . $kandidat->image_calon);

            if (file_exists($imagePath)) {
                unlink($imagePath);
            }

            return redirect()->route('kandidat')->with('success', 'Kandidat berhasil dihapus!');
        } catch (\Exception $e) {
            // return redirect()->back()->with('error', $e->getMessage());
            return redirect()->back()->with('error', 'Gagal menghapus! Silakan hapus data terkait di tabel votes terlebih dahulu.');
        }
    }

    public function scoreAkhir()
    {
        $kandidat = KandidatModel::withCount('votes')->get();

        return view('dashboard.score', compact('kandidat'));
    }
    // KandidatController.php
    public function show($id)
    {
        $kandidat = KandidatModel::with('votes.user')->findOrFail($id); // Fetch candidate with voters
        $voters = $kandidat->votes->map->user;
        $totalVoters = $voters->count();

        return view('dashboard/components.show', compact('kandidat', 'voters', 'totalVoters'));
    }

}