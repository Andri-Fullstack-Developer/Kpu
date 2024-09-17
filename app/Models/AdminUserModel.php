<?php

namespace App\Models;

use Illuminate\Database\Eloquent\Factories\HasFactory;
use Illuminate\Database\Eloquent\Model;

use Illuminate\Foundation\Auth\User as Authenticatable;

class AdminUserModel extends Authenticatable
{
    use HasFactory;

    protected $table = 'admin_users';
    protected $fillable = [
        'name_admin',
        'username_admin',
        'password_admin',
        'role_admin',
    ];

    protected $hidden = [
        'password_admin',
    ];

    public $timestamps = true;

    public function getAuthPassword()
    {
        return $this->password_admin;
    }
}
