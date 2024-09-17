@extends('dashboard/layouts/main')

@section('content')
    <div class="w-4/12 absolute -top-4 z-10 flex items-end right-0">
        @include('dashboard/components.alerd')
    </div>
    <h1 class="text-3xl text-black pb-6">List User</h1>

    <!-- Form Penghapusan -->
    <form action="{{ route('delete') }}" method="POST" id="deleteForm">
        @csrf
        @method('DELETE')
        <button type="submit" class="bg-red-500 text-white btn btn-sm mb-4 hidden" id="deleteButton">Delete</button>
    </form>

    <!-- User Table -->
    <div class="overflow-x-auto">
        <table id="list" class="min-w-full text-xs rounded">
            <thead class="rounded-t-lg text-white bg-gray-800">
                <tr class="text-right">
                    <th class="px-4 py-2 text-sm">
                        <div class="flex gap-3">
                            <span class="label-text text-sm text-white">All</span>
                            <input type="checkbox" id="checkAll"
                                class="h-4 w-4 rounded border-gray-300 bg-white text-teal-600" />
                        </div>
                    </th>
                    <th class="px-4 border py-2 text-sm">Name</th>
                    <th class="px-4 border py-2 text-sm">Username</th>
                    <th class="px-4 border py-2 text-sm">Kelas</th>
                    <th class="px-4 border py-2 text-sm">Status</th>
                    <th class="px-4 border py-2 text-sm">Action</th>
                </tr>
            </thead>
            <tbody>
                @foreach ($users as $user)
                    <tr class="text-black">
                        <td class="border px-4 py-2">
                            <input type="checkbox" name="selected_users[]" value="{{ $user->id }}"
                                class="checkItem h-4 w-4 rounded border-gray-300 bg-white text-teal-600">
                        </td>
                        <td class="border px-4 py-2 text-sm">{{ $user->name }}</td>
                        <td class="border px-4 py-2 text-sm">{{ $user->username }}</td>
                        <td class="border px-4 py-2 text-sm">{{ $user->kelas }}</td>
                        <td class="border px-4 py-2 text-sm">{{ $user->status }}</td>
                        <td class="border px-4 py-2 text-sm text-center">
                            <button type="button"
                                onclick="document.getElementById('editUser{{ $user->id }}').showModal()">
                                <i class="fa-duotone fa-solid fa-file-pen"></i>
                            </button>
                        </td>
                    </tr>

                    <!-- Modal Edit User -->
                    <dialog id="editUser{{ $user->id }}" class="modal">
                        @include('dashboard/components/edit-list')
                    </dialog>
                @endforeach
            </tbody>
        </table>
    </div>

    <!-- JavaScript -->
    <script>
        document.addEventListener('DOMContentLoaded', function() {
            $('#list').DataTable({});

            function toggleDeleteButton() {
                const checkItems = document.querySelectorAll('.checkItem');
                const deleteButton = document.getElementById('deleteButton');
                let isAnyChecked = Array.from(checkItems).some(item => item.checked);

                deleteButton.classList.toggle('hidden', !isAnyChecked);
            }

            // Handle "Check All" functionality
            document.getElementById('checkAll').addEventListener('click', function(e) {
                let checkItems = document.querySelectorAll('.checkItem');
                checkItems.forEach(item => item.checked = e.target.checked);
                toggleDeleteButton();
            });

            // Handle individual checkbox click events
            document.querySelectorAll('.checkItem').forEach(item => {
                item.addEventListener('click', function() {
                    toggleDeleteButton();
                });
            });

            // Handle form submission with selected IDs
            document.getElementById('deleteButton').addEventListener('click', function() {
                const selectedIds = Array.from(document.querySelectorAll('.checkItem:checked')).map(item =>
                    item.value);
                const form = document.getElementById('deleteForm');

                // Remove any previously added hidden inputs
                form.querySelectorAll('input[name="selected_users[]"]').forEach(input => input.remove());

                // Add selected IDs as a hidden input field
                selectedIds.forEach(id => {
                    const input = document.createElement('input');
                    input.type = 'hidden';
                    input.name = 'selected_users[]';
                    input.value = id;
                    form.appendChild(input);
                });

                // Submit the form
                form.submit();
            });
        });
    </script>
@endsection
