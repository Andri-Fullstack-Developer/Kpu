@extends('dashboard/layouts/main')

@section('content')
    <div class="w-4/12 absolute -top-4 z-10 flex items-end right-0">
        @include('dashboard/components.alerd')
    </div>
    <form action="{{ route('delete.votes') }}" method="POST" id="deleteForm">
        @csrf
        @method('DELETE')
        <button type="submit" class="bg-red-500 text-white btn btn-sm mb-4 hidden">Delete</button>
        <div class="overflow-x-auto">
            <table id="votes" class="items-center w-full bg-transparent border-collapse">
                <thead>
                    <tr>
                        <th class="px-4 py-2 text-sm bg-gray-50 text-gray-700">
                            <div class="flex gap-3">
                                <span class="label-text text-sm">All</span>
                                <input type="checkbox" id="checkAll" class="checkbox-info" />
                            </div>
                        </th>
                        <th
                            class="px-4 bg-gray-50 text-gray-700 align-middle py-3 text-xs font-semibold text-left uppercase border-l-0 border-r-0 whitespace-nowrap">
                            #</th>
                        <th
                            class="px-4 bg-gray-50 text-gray-700 align-middle py-3 text-xs font-semibold text-left uppercase border-l-0 border-r-0 whitespace-nowrap">
                            Name Calon</th>
                        <th
                            class="px-4 bg-gray-50 text-gray-700 align-middle py-3 text-xs font-semibold text-left uppercase border-l-0 border-r-0 whitespace-nowrap">
                            Kelas</th>
                        <th
                            class="px-4 bg-gray-50 text-gray-700 align-middle py-3 text-xs font-semibold text-left uppercase border-l-0 border-r-0 whitespace-nowrap">
                            Name</th>
                        <th
                            class="px-4 bg-gray-50 text-gray-700 align-middle py-3 text-xs font-semibold text-left uppercase border-l-0 border-r-0 whitespace-nowrap">
                            Username</th>
                    </tr>
                </thead>
                <tbody class="divide-y divide-gray-100">
                    @foreach ($vote as $i => $vote)
                        <tr class="text-gray-500">
                            <td class="border px-4 py-2">
                                <input type="checkbox" name="selected_votes[]" value="{{ $vote->id }}"
                                    class="checkItem">
                            </td>
                            <th class="border-t-0 px-4 align-middle text-sm font-normal whitespace-nowrap p-4 text-left">
                                {{ $i + 1 }}</th>
                            <td
                                class="border-t-0 px-4 align-middle text-xs font-medium text-gray-900 whitespace-nowrap p-4">
                                {{ $vote->kandidat->name_calon }}</td>
                            <td
                                class="border-t-0 px-4 align-middle text-xs font-medium text-gray-900 whitespace-nowrap p-4">
                                {{ $vote->user->kelas }}</td>
                            <td
                                class="border-t-0 px-4 align-middle text-xs font-medium text-gray-900 whitespace-nowrap p-4">
                                {{ $vote->user->name }}</td>
                            <td
                                class="border-t-0 px-4 align-middle text-xs font-medium text-gray-900 whitespace-nowrap p-4">
                                {{ $vote->user->username }}</td>
                        </tr>
                    @endforeach
                </tbody>
            </table>
        </div>
    </form>

    <script>
        $(document).ready(function() {
            $('#votes').DataTable();

            const toggleDeleteVotes = () => {
                const checkItems = document.querySelectorAll('.checkItem');
                const deleteVotes = document.querySelector('button[type="submit"]');
                const isAnyChecked = Array.from(checkItems).some(item => item.checked);

                deleteVotes.classList.toggle('hidden', !isAnyChecked);
            };

            $('#checkAll').on('click', function() {
                const checkItems = document.querySelectorAll('.checkItem');
                checkItems.forEach(item => item.checked = this.checked);
                toggleDeleteVotes();
            });

            document.querySelectorAll('.checkItem').forEach(item => {
                item.addEventListener('change', toggleDeleteVotes);
            });
        });
    </script>
@endsection
