@extends('dashboard/layouts/main')

@section('content')
    <div class="grid md:grid-cols-3 grid-cols-2 justify-center md:gap-4 gap-2 md:px-3">
        @foreach ($kandidat as $item)
            <a href="{{ route('score.show', $item->id) }}">
                <div
                    class="max-w-xs container bg-white rounded-xl shadow-lg transform transition duration-500 hover:scale-105 hover:shadow-2xl">
                    <div>
                        <h1
                            class="text-2xl mt-2 ml-4 font-bold text-gray-800 cursor-pointer hover:text-gray-900 transition duration-100">
                            {{ $item->name_calon }}</h1>
                        <p class="ml-4 mt-1 mb-2 text-gray-700 hover:underline cursor-pointer">
                            {{ $item->kelas_calon }}</p>
                    </div>
                    <img class="h-80 w-full object-cover " src="{{ asset('images/' . $item->image_calon) }}" alt="" />
                    <h1 class="text-lg font-semibold text-gray-800">
                        Score:
                        @if ($item->votes_count > 0)
                            {{ number_format($item->votes_count) }}
                        @else
                            No votes yet
                        @endif
                    </h1>
                </div>
            </a>
        @endforeach
    </div>
@endsection
