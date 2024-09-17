   <!-- Desktop Header -->
   <header class="w-full items-center bg-white py-2 px-6 hidden sm:flex">
       <div class="w-1/2">
           <div class="w-10 h-10">
               <img src="{{ asset('assets/images/smpn1.png') }}">
           </div>
       </div>
       <div class="relative w-1/2 flex justify-end">
           <form action="{{ route('logout.admin') }}" method="POST">
               @csrf
               <button class="block px-4 py-2 border border-2 rounded-lg account-link hover:text-white">Sign Out</button>
           </form>
       </div>
   </header>

   <!-- Mobile Header & Nav -->
   <header x-data="{ isOpen: false }" class="w-full bg-sidebar py-5 px-6 sm:hidden">
       <div class="flex items-center justify-between">
           <a href="index.html"
               class="text-white text-lg  font-semibold uppercase hover:text-gray-300">{{ $role }}</a>
           <button @click="isOpen = !isOpen" class="text-white text-3xl focus:outline-none">
               <i x-show="!isOpen" class="fas fa-bars"></i>
               <i x-show="isOpen" class="fas fa-times"></i>
           </button>
       </div>

       <!-- Dropdown Nav -->
       <nav :class="isOpen ? 'flex' : 'hidden'" class="flex flex-col pt-4">
           <a href="{{ route('dashboard') }}"
               class="flex items-center {{ request()->routeIs('dashboard') ? 'active-nav-link' : '' }} text-white py-2 pl-4 nav-item">
               <i class="fa-duotone fa-solid fa-gauge mr-3"></i>
               Dashboard
           </a>
           <a href="{{ route('tambah-user') }}"
               class="flex items-center {{ request()->routeIs('tambah-user') ? 'active-nav-link' : '' }} text-white opacity-75 hover:opacity-100 py-2 pl-4 nav-item">
               <i class="fa-duotone fa-solid fa-users-medical mr-3"></i>
               Tambah User
           </a>
           <a href="{{ route('listUsers') }}"
               class="flex items-center {{ request()->routeIs('listUsers') ? 'active-nav-link' : '' }} text-white opacity-75 hover:opacity-100 py-2 pl-4 nav-item">
               <i class="fa-duotone fa-solid fa-grid mr-3"></i>
               List Users
           </a>
           <a href="{{ route('tambah-admin') }}"
               class="flex items-center {{ request()->routeIs('tambah-admin') ? 'active-nav-link' : '' }} text-white opacity-75 hover:opacity-100 py-2 pl-4 nav-item">
               <i class="fa-duotone fa-solid fa-user-plus mr-3"></i>
               Tambah Admin
           </a>
           <a href="{{ route('kandidat') }}"
               class="flex items-center {{ request()->routeIs('kandidat') ? 'active-nav-link' : '' }} text-white opacity-75 hover:opacity-100 py-2 pl-4 nav-item">
               <i class="fa-sharp-duotone fa-solid fa-user-plus mr-3"></i>
               Tambah Kandidat
           </a>
           <a href="{{ route('scoreAkhir') }}"
               class="flex items-center {{ request()->routeIs('scoreAkhir') ? 'active-nav-link' : '' }} text-white opacity-75 hover:opacity-100 py-2 pl-4 nav-item">
               <i class="fa-duotone fa-solid fa-star-exclamation mr-3"></i>
               Score
           </a>
           <a href="{{ route('votes') }}"
               class="flex items-center {{ request()->routeIs('votes') ? 'active-nav-link' : '' }} text-white opacity-75 hover:opacity-100 py-2 pl-4 nav-item">
               <i class="fa-duotone fa-solid fa-check-to-slot mr-3"></i>
               Votes
           </a>
           <form action="{{ route('logout.admin') }}" method="POST">
               @csrf
               <button class="w-full">
                   <a href="#"
                       class="flex items-center text-white opacity-75 hover:opacity-100 py-2 pl-4 nav-item">
                       <i class="fas fa-sign-out-alt mr-3"></i>
                       Sign Out
                   </a>
               </button>
           </form>
       </nav>
   </header>
