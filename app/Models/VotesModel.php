<?php

namespace App\Models;

use Illuminate\Database\Eloquent\Factories\HasFactory;
use Illuminate\Database\Eloquent\Model;

class VotesModel extends Model
{
    use HasFactory;

    protected $table = 'votes';

    protected $fillable = [
        'user_id',
        'candidate_id',
        'vote_time',
    ];

    public $timestamps = true;

//     public function users()
//     {
//         return $this->belongsTo(User::class, 'user_id');
//     }
//
//     public function candidate()
//     {
//         return $this->belongsTo(KandidatModel::class, 'candidate_id');
//     }
//
//     public function user()
//     {
//         return $this->belongsTo(User::class);
//     }
//
//     // Define the relationship to the candidate
//     public function kandidat()
//     {
//         return $this->belongsTo(KandidatModel::class);
//     }

    public function user()
    {
        return $this->belongsTo(User::class, 'user_id');
    }

    public function kandidat()
    {
        return $this->belongsTo(KandidatModel::class, 'candidate_id');
    }
}