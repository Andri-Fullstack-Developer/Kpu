<?php

namespace App\Models;

use Illuminate\Database\Eloquent\Factories\HasFactory;
use Illuminate\Database\Eloquent\Model;

class KandidatModel extends Model
{
    use HasFactory;

    protected $table = 'kandidats';

    protected $fillable = [
        'name_calon',
        'kelas_calon',
        'visi_calon',
        'misi_calon',
        'image_calon',
    ];

    public $timestamps = true;

    public function votes()
    {
        return $this->hasMany(VotesModel::class, 'candidate_id');
    }
    public function voters()
    {
        return $this->hasMany(VotesModel::class);
    }
}