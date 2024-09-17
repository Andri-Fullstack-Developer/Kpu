<div class="p-4">
    <a href=""
        class="text-blackF bg-[#EDECE2] rounded px-4 py-1 text-xl font-semibold uppercase hover:text-gray-300">{{ $role }}</a>

</div>
<nav class="text-white text-base font-semibold pt-3">
    <a href="{{ route('dashboard') }}"
        class="flex items-center {{ request()->routeIs('dashboard') ? 'active-nav-link' : '' }} text-white py-4 pl-6 nav-item">
        <i class="fa-duotone fa-solid fa-gauge mr-3"></i>
        Dashboard
    </a>
    <a href="{{ route('tambah-user') }}"
        class="flex items-center {{ request()->routeIs('tambah-user') ? 'active-nav-link' : '' }} text-white opacity-75 hover:opacity-100 py-4 pl-6 nav-item">
        <i class="fa-duotone fa-solid fa-users-medical mr-3"></i>
        Tambah User
    </a>
    <a href="{{ route('listUsers') }}"
        class="flex items-center {{ request()->routeIs('listUsers') ? 'active-nav-link' : '' }} text-white opacity-75 hover:opacity-100 py-4 pl-6 nav-item">
        <i class="fa-duotone fa-solid fa-grid mr-3"></i>
        List Users
    </a>
    <a href="{{ route('tambah-admin') }}"
        class="flex items-center {{ request()->routeIs('tambah-admin') ? 'active-nav-link' : '' }} text-white opacity-75 hover:opacity-100 py-4 pl-6 nav-item">
        <i class="fa-duotone fa-solid fa-user-plus mr-3"></i>
        Tambah Admin
    </a>
    <a href="{{ route('kandidat') }}"
        class="flex items-center {{ request()->routeIs('kandidat') ? 'active-nav-link' : '' }} text-white opacity-75 hover:opacity-100 py-4 pl-6 nav-item">
        <i class="fa-sharp-duotone fa-solid fa-user-plus mr-3"></i>
        Tambah Kandidat
    </a>
    <a href="{{ route('scoreAkhir') }}"
        class="flex items-center {{ request()->routeIs('scoreAkhir') ? 'active-nav-link' : '' }} text-white opacity-75 hover:opacity-100 py-4 pl-6 nav-item">
        <i class="fa-duotone fa-solid fa-star-exclamation mr-3"></i>
        Score
    </a>
    <a href="{{ route('votes') }}"
        class="flex items-center {{ request()->routeIs('votes') ? 'active-nav-link' : '' }} text-white opacity-75 hover:opacity-100 py-4 pl-6 nav-item">
        <i class="fa-duotone fa-solid fa-check-to-slot mr-3"></i>
        Votes
    </a>
</nav>
