@extends('dashboard/layouts/main')

@section('content')
    <h1 class="text-3xl text-black pb-6">Tambah User</h1>
    <div class="w-max md:hidden flex gap-6">
        <button class="px-10 py-1 bg-blue-400 rounded" id="show-user-data">User Data</button>
        <button class="px-10 py-1 bg-blue-400 rounded" id="show-user-exsel">User Exsel</button>
    </div>
    <div class="w-4/12 absolute -top-4 z-10 flex items-end right-0">
        @include('dashboard/components.alerd')
    </div>
    <div class="grid md:grid-cols-2 grid-col-1 gap-2">
        <div class="w-full" id="user-data">
            <form action="{{ route('tambah') }}" method="POST">
                @csrf
                <div class="mb-4">
                    <label class="block text-gray-700 font-bold mb-2" for="name">
                        Nama
                    </label>
                    <input
                        class="shadow appearance-none border rounded bg-gray-100 w-full py-2 px-3 text-gray-700 leading-tight focus:outline-none focus:shadow-outline"
                        id="name" type="text" placeholder="Your Name..." name="name">
                </div>
                <div class="mb-4">
                    <label class="block text-gray-700 font-bold mb-2" for="nis">
                        NIS
                    </label>
                    <input
                        class="shadow appearance-none border rounded bg-gray-100 w-full py-2 px-3 text-gray-700 leading-tight focus:outline-none focus:shadow-outline"
                        id="nis" type="text" placeholder="Your NIS..." name="username">
                </div>
                <div class="mb-4">
                    <label class="block text-gray-700 font-bold mb-2" for="kelas">
                        Kelas
                    </label>
                    <input
                        class="shadow appearance-none border rounded bg-gray-100 w-full py-2 px-3 text-gray-700 leading-tight focus:outline-none focus:shadow-outline"
                        id="kelas" type="text" placeholder="Your Kelas..." name="kelas">
                </div>
                <div class="mb-4">
                    <label class="block text-gray-700 font-bold mb-2" for="password">
                        Password
                    </label>
                    <input
                        class="shadow appearance-none border rounded bg-gray-100 w-full py-2 px-3 text-gray-700 leading-tight focus:outline-none focus:shadow-outline"
                        id="password" type="password" placeholder="Your Password..." name="password">
                </div>
                <button type="submit" class="btn btn-active bg-[#56AEF7] text-white border-none">Submit</button>
            </form>
        </div>
        <div class="w-full hidden md:block" id="user-exsel">
            <form action="{{ route('importData') }}" method="POST" enctype="multipart/form-data">
                @csrf
                <div class="mb-4">
                    <label class="block text-gray-700 font-bold mb-2" for="photo-dropbox">
                        Upload
                    </label>
                    <label
                        class="flex  cursor-pointer appearance-none justify-center rounded-md border border-dashed border-gray-300 bg-white px-3 py-6 text-sm transition hover:border-gray-400 focus:border-solid focus:border-blue-600 focus:outline-none focus:ring-1 focus:ring-blue-600 disabled:cursor-not-allowed disabled:bg-gray-200 disabled:opacity-75"
                        tabindex="0">
                        <span for="photo-dropbox" class="flex items-center space-x-2">
                            <i class="fa-duotone fa-solid fa-cloud-arrow-up"></i>
                            <span class="text-xs font-medium text-gray-600">
                                Drop files to Attach, or
                                <span class="text-blue-600 underline">browse</span>
                            </span>
                        </span>
                        <input id="photo-dropbox" type="file" class="sr-only" name="upload" accept=".xls,.xlsx" />
                    </label>
                    <div id="file-message" class=" bg-red-300 text-white hidden mt-2 rounded-lg">
                        <p id="file-name" class="text-sm text-gray-600 py-2 px-4 "></p>
                    </div>
                    <div class="mt-1 p-2 bg-white rounded">
                        <p>
                            <span class="text-gray-900">*Format yang diperbolehkan: Excel (.xls,.xlsx)</span>
                            <a href="{{ asset('assets/files/data_login.xlsx') }}" download="data_login.xlsx"
                                class="text-red-500 underline">download</a>
                        </p>
                    </div>
                </div>

                <button type="submit" class="btn btn-active bg-[#56AEF7] text-white border-none">Save</button>
            </form>
        </div>
    </div>
    <script>
        $(document).ready(function() {
            $('#photo-dropbox').change(function() {
                const file = $(this).val();
                const fileName = file.split('\\').pop();
                const extension = file.split('.').pop().toLowerCase();
                const allowedExtensions = ['xls', 'xlsx'];

                // Tampilkan pesan berdasarkan validasi file
                if (allowedExtensions.includes(extension)) {
                    $('#file-name').text(fileName);
                    $('#file-message').removeClass('bg-red-300').addClass('bg-green-300').removeClass(
                        'hidden');
                } else {
                    $('#file-name').text('Invalid file type. Please upload an Excel file.');
                    $('#file-message').removeClass('hidden').addClass(
                        'bg-red-300'); // Tampilkan pesan kesalahan
                    $(this).val(''); // Kosongkan input jika file tidak valid
                }
            });

            @if (session('error'))
                $('#file-message').removeClass('hidden bg-green-300').addClass('bg-red-300');
                $('#file-name').text("{{ session('error') }}");
            @elseif (session('success'))
                $('#file-message').removeClass('hidden bg-red-300').addClass('bg-green-300');
                $('#file-name').text("{{ session('success') }}");
            @endif

            function setActiveButton(buttonId) {
                $('#show-user-data, #show-user-exsel').removeClass('bg-blue-700 text-white').addClass(
                    'bg-blue-400 text-black');
                $(buttonId).removeClass('bg-blue-400 text-black').addClass('bg-blue-700 text-white');
            }

            $('#show-user-data').on('click', function() {
                $('#user-data').show();
                $('#user-exsel').hide();
                setActiveButton('#show-user-data');
            });

            $('#show-user-exsel').on('click', function() {
                $('#user-data').hide();
                $('#user-exsel').show();
                setActiveButton('#show-user-exsel');
            });

            $(window).resize(function() {
                if ($(window).width() >= 768) {
                    $('#user-data').show();
                    $('#user-exsel').show();
                    $('#show-user-data, #show-user-exsel').removeClass('bg-blue-700 text-white').addClass(
                        'bg-blue-400 text-black');
                } else {
                    $('#user-data').show();
                    $('#user-exsel').hide();
                    setActiveButton('#show-user-data');
                }
            }).resize(); // Trigger resize event on page load
        });
    </script>
@endsection


{{-- <div class="mb-4">
                    <label class="block text-gray-700 font-bold mb-2" for="name">
                        Import Data Exsel
                    </label>
                    <input type="file" id="upload" name="upload"
                        class="shadow appearance-none border rounded w-full py-2 px-3 text-gray-700 leading-tight focus:outline-none focus:shadow-outline file-input file-input-bordered"
                        required />
                    @error('upload')
                        <p class="text-red-500 text-xs mt-2">{{ $message }}</p>
                    @enderror
                </div> --}}
