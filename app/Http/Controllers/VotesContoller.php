<?php

namespace App\Http\Controllers;

use App\Events\VoteUpdated;
use App\Models\KandidatModel;
use App\Models\User;
use App\Models\VotesModel;
use Illuminate\Http\Request;
use Illuminate\Support\Facades\Auth;

class VotesContoller extends Controller
{
    public function store(Request $request)
    {
        $request->validate([
            'candidate_id' => 'required|exists:kandidats,id',
        ]);

        $user = Auth::user();

        if ($user) {
            date_default_timezone_set('Asia/Jakarta');
            $vote = now();

            VotesModel::create([
                'user_id' => $user->id,
                'candidate_id' => $request->candidate_id,
                'vote_time' => $vote,
            ]);

            $user->status = 'tidak aktif';
            $user->save();

            Auth::logout();
            return redirect()->route('thanksKpu');

            // return redirect()->back()->with('success', 'Your vote has been recorded.');
        }

        return redirect()->back()->with('error', 'User not found.');

    }

    public function castVote(Request $request)
    {
        // Validasi input
        $request->validate([
            'candidate_id' => 'required|exists:kandidats,id',
        ]);

        $candidateId = $request->input('candidate_id');

        // Misalkan Anda melakukan pencatatan vote di sini
        // Contoh: updating the vote count in the database

        // Ambil jumlah vote terbaru
        $votesCount = KandidatModel::find($candidateId)->votes_count;

        // Data untuk dikirimkan ke front-end
        $voteData = [
            'candidate_id' => $candidateId,
            'votes_count' => $votesCount,
        ];

        // Kirimkan event ke front-end
        event(new VoteUpdated($voteData));

        // Respons jika diperlukan
        return response()->json(['success' => true]);
    }

    public function deleteVotes(Request $req)
    {
        $voteIds = $req->input('selected_votes', []);

        $votes = VotesModel::whereIn('id', $voteIds)->get();

        foreach ($votes as $vote) {
            $user = $vote->user;
            $user->status = 'aktif';
            $user->save();
        }

        // Delete the votes
        VotesModel::whereIn('id', $voteIds)->delete();

        return redirect()->route('votes')->with('success', 'Votes deleted and users updated successfully.');

    }

}