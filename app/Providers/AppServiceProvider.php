<?php

namespace App\Providers;

use Illuminate\Support\Facades\Auth;
use Illuminate\Support\Facades\View;
use Illuminate\Support\ServiceProvider;

class AppServiceProvider extends ServiceProvider
{
    /**
     * Register any application services.
     */
    public function register(): void
    {
        //
    }

    /**
     * Bootstrap any application services.
     */
    public function boot(): void
    {
        View::composer('*', function ($view) {
            $admin = Auth::guard('admin')->user();
            $view->with('role', $admin ? $admin->role_admin : null);
            // $view->with('adminRole', $admin ? $admin->role_admin : null);
        });
    }
}
