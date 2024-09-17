@extends('dashboard/layouts/main')

@section('content')
    <h1 class="text-3xl text-black pb-6">Tambah Admin</h1>
    <div class="w-4/12 absolute -top-4 z-10 flex items-end right-0">
        @include('dashboard/components.alerd')
    </div>
    <div class="w-max md:hidden flex gap-6">
        <button class="px-10 py-1 bg-blue-400 rounded" id="show-tambah-admin">Tambah Admin</button>
        <button class="px-10 py-1 bg-blue-400 rounded" id="show-tabel-admin">Tabel Admin</button>
    </div>
    <div class="grid grid-cols-1 md:grid-cols-2 gap-2">
        <div id="tambah-admin" class="">
            <form action="{{ route('createAdmin') }}" method="POST">
                @csrf
                <div class="mb-4 ">
                    <label class="block text-gray-700 font-bold mb-2" for="name">
                        Nama Admin
                    </label>
                    <input
                        class="shadow appearance-none border rounded w-full py-2 px-3 text-gray-700 leading-tight focus:outline-none focus:shadow-outline bg-white"
                        id="name" type="text" placeholder="Your Nama" name="name-admin">
                </div>
                <div class="mb-4 ">
                    <label class="block text-gray-700 font-bold mb-2" for="nis-admin">
                        NIS Admin
                    </label>
                    <input
                        class="shadow appearance-none border rounded w-full py-2 px-3 text-gray-700 leading-tight focus:outline-none focus:shadow-outline bg-white"
                        id="nis-admin" type="text" placeholder="Your Nis Admin" name="username-admin">
                </div>
                <div class="mb-4 ">
                    <label class="block text-gray-700 font-bold mb-2" for="password-admin">
                        Password Admin
                    </label>
                    <input
                        class="shadow appearance-none border rounded w-full py-2 px-3 text-gray-700 leading-tight focus:outline-none focus:shadow-outline bg-white"
                        id="password-admin" type="text" placeholder="Your password Admin" name="password-admin">
                </div>
                <button type="submit" class="btn btn-primary text-white">Save</button>
            </form>
        </div>
        <div id="tabel-admin" class="hidden md:block">
            <div class="container p-2 mx-auto rounded-md sm:p-4 dark:text-gray-800 dark:bg-gray-50">
                <h2 class="mb-3 text-2xl font-semibold leading-tight">Table user</h2>
                <div class="overflow-x-auto">
                    <table class="min-w-full text-xs rounded">
                        <thead class="rounded-t-lg text-black bg-gray-100">
                            <tr class="text-right">
                                <th class="p-3 text-left">#</th>
                                <th class="p-3 text-left">Name</th>
                                <th class="p-3">Username</th>
                                <th class="p-3">Role</th>
                                <th class="p-3">Action</th>
                            </tr>
                        </thead>
                        <tbody>
                            @foreach ($list_admin as $i => $item)
                                <tr class="text-right border border-opacity-20 bg-gray-300 border-black">
                                    <td class="px-3 py-2 text-left">
                                        <span>{{ $i + 1 }}</span>
                                    </td>
                                    <td class="px-3 py-2 text-left">
                                        <span>{{ $item->name_admin }}</span>
                                    </td>
                                    <td class="px-3 py-2">
                                        <span>{{ $item->username_admin }}</span>
                                    </td>
                                    <td class="px-3 py-2">
                                        <span>{{ $item->role_admin }}</span>
                                    </td>
                                    <td class="px-3 py-2">
                                        <span>
                                            <button class=""
                                                onclick="document.getElementById('show_Delti{{ $item->id }}').showModal()">
                                                <i class="fa-duotone fa-solid fa-grid-2"></i>
                                            </button>
                                        </span>
                                        @include('dashboard/components/modal-admin')
                                    </td>
                                </tr>
                            @endforeach
                        </tbody>
                    </table>
                </div>
            </div>
        </div>
    </div>
    <script>
        function closeModal(modalId) {
            $('#' + modalId)[0].close();
        }

        $(document).ready(function() {
            // Fungsi untuk mengaktifkan tombol yang diklik
            function setActiveButton(buttonId) {
                $('#show-tambah-admin, #show-tabel-admin').removeClass('bg-blue-700 text-white');
                $('#show-tambah-admin, #show-tabel-admin').addClass('bg-blue-400 text-black');

                // Tambahkan kelas 'bg-red-700' ke tombol yang diklik
                $(buttonId).removeClass('bg-blue-400 text-black').addClass('bg-blue-700 text-white');
            }

            $(window).resize(function() {
                if ($(window).width() >= 768) {
                    // Tampilkan kedua div di desktop (>= 768px)
                    $('#tambah-admin').show();
                    $('#tabel-admin').show();
                    // Hapus kelas 'bg-blue-700' dari semua tombol
                    $('#show-tambah-admin, #show-tabel-admin').removeClass('bg-blue-700 text-white')
                        .addClass('bg-blue-400 text-black');
                } else {
                    // Pada mobile, sembunyikan "Tabel Admin" dan tampilkan "Tambah Admin" di awal
                    $('#tambah-admin').show();
                    $('#tabel-admin').hide();
                    // Aktifkan tombol "Tambah Admin" pada mobile
                    setActiveButton('#show-tambah-admin');
                }
            }).resize();

            // Ketika tombol "Tambah Admin" diklik
            $('#show-tambah-admin').on('click', function() {
                $('#tabel-admin').hide(); // Sembunyikan "Tabel Admin"
                $('#tambah-admin').show(); // Tampilkan "Tambah Admin"
                setActiveButton('#show-tambah-admin'); // Set tombol ini aktif
            });

            // Ketika tombol "Tabel Admin" diklik
            $('#show-tabel-admin').on('click', function() {
                $('#tambah-admin').hide(); // Sembunyikan "Tambah Admin"
                $('#tabel-admin').show(); // Tampilkan "Tabel Admin"
                setActiveButton('#show-tabel-admin'); // Set tombol ini aktif
            });
        });
    </script>
@endsection
