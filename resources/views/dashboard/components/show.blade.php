{{-- resources/views/kandidat/show.blade.php --}}
@extends('dashboard.layouts.main')

@section('content')
    <div class="mb-4">
        <a href="{{ url()->previous() }}" class="border border-2 rounded-md px-4 py-3 bg-blue-700 text-white">
            <i class="fa-duotone fa-solid fa-hand-back-point-left me-2 "></i> Back
        </a>
    </div>
    <div class="grid md:grid-cols-3 grid-cols-1 gap-4">
        <div
            class="md:max-w-xs w-full container bg-white rounded-xl shadow-lg transform transition duration-500 hover:scale-52 hover:shadow-2xl">
            <div>
                <h1
                    class="text-2xl mt-2 ml-4 font-bold text-gray-800 cursor-pointer hover:text-gray-900 transition duration-100">
                    {{ $kandidat->name_calon }}</h1>
                <p class="ml-4 mt-1 mb-2 text-gray-700 hover:underline cursor-pointer">
                    {{ $kandidat->kelas_calon }}</p>
            </div>
            <img class="w-full cursor-pointer" src="{{ asset('images/' . $kandidat->image_calon) }}" alt="" />

        </div>
        <div class="col-span-2 p-2 block w-full overflow-x-auto mb-5 border">
            <table id="show" class="items-center w-full bg-transparent border-collapse">
                <thead>
                    <tr>
                        <th
                            class="px-4 bg-gray-50 text-gray-700 align-middle py-3 text-xs font-semibold text-left uppercase border-l-0 border-r-0 whitespace-nowrap">
                            #</th>
                        <th
                            class="px-4 bg-gray-50 text-gray-700 align-middle py-3 text-xs font-semibold text-left uppercase border-l-0 border-r-0 whitespace-nowrap">
                            Name</th>

                        <th
                            class="px-4 bg-gray-50 text-gray-700 align-middle py-3 text-xs font-semibold text-left uppercase border-l-0 border-r-0 whitespace-nowrap">
                            kelas</th>
                        <th
                            class="px-4 bg-gray-50 text-gray-700 align-middle py-3 text-xs font-semibold text-left uppercase border-l-0 border-r-0 whitespace-nowrap">
                            Username</th>
                    </tr>
                </thead>
                <tbody class="divide-y divide-gray-100">
                    @foreach ($voters as $i => $user)
                        <tr class="text-gray-500">
                            <th class="border-t-0 px-4 align-middle text-sm font-normal whitespace-nowrap p-4 text-left">
                                {{ $i + 1 }}</th>
                            <td
                                class="border-t-0 px-4 align-middle text-xs font-medium text-gray-900 whitespace-nowrap p-4">
                                {{ $user->name }}</td>
                            <td
                                class="border-t-0 px-4 align-middle text-xs font-medium text-gray-900 whitespace-nowrap p-4">
                                {{ $user->kelas }}</td>
                            <td
                                class="border-t-0 px-4 align-middle text-xs font-medium text-gray-900 whitespace-nowrap p-4">
                                {{ $user->username }}</td>
                        </tr>
                    @endforeach
                </tbody>
            </table>

            <button class="px-6 py-3 rounded-lg bg-info text-white">Voters:
                <span>{{ $totalVoters }}</span></button>
        </div>
    </div>
    <script>
        $(document).ready(function() {
            $('#show').DataTable({

            });
        });
    </script>
@endsection
