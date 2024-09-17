@extends('dashboard/layouts/main')

@section('content')
    <div class="w-4/12 absolute -top-4 z-10 flex justify-end right-0">
        @include('dashboard/components.alerd')
    </div>
    <div class="relative h-screen">
        <h1 class="text-3xl text-black pb-6">Tambah Kandidat</h1>
        <button class="btn bg-blue-400 text-white btn-outline mb-4" onclick="new_kandidat.showModal()">New Kandidat</button>
        <div
            class="w-fit mx-auto grid grid-cols-1 lg:grid-cols-3 md:grid-cols-2 justify-items-center justify-center gap-y-20 gap-10 mt-10 mb-5 relative ">
            @foreach ($kandidat as $item)
                <div class="w-72 bg-white shadow-md rounded-xl duration-500 hover:scale-105 hover:shadow-xl relative">
                    <a href="#">
                        <img src="{{ asset('images/' . $item->image_calon) }}" alt="Product"
                            class="h-80 w-72 object-cover rounded-t-xl" />
                        <div class="px-4 py-3 w-72 flex justify-between text-white">
                            <button onclick="document.getElementById('show{{ $item->id }}').showModal()"
                                class="px-5 py-1 border-2 rounded bg-blue-500">
                                <i class="fa-duotone fa-solid fa-eye"></i>
                            </button>
                            <button onclick="document.getElementById('edit{{ $item->id }}').showModal()"
                                class="px-5 py-1 border-2 rounded bg-blue-500"><i
                                    class="fa-duotone fa-solid fa-pen-to-square"></i></button>
                            <button onclick="document.getElementById('delete{{ $item->id }}').showModal()"
                                class="px-5 py-1 border-2 rounded bg-blue-500">
                                <i class="fa-duotone fa-solid fa-trash"></i>
                            </button>

                        </div>
                    </a>
                </div>
                <!-- You can open the modal using ID.showModal() method -->

                <!-- Delete -->
                <dialog id="delete{{ $item->id }}" class="modal">
                    <div class="modal-box bg-white">
                        <form method="dialog">
                            <button class="btn btn-sm btn-circle btn-ghost absolute right-2 top-2 text-3xl"><i
                                    class="fa-duotone fa-solid fa-circle-xmark text-black"></i></button>
                        </form>
                        <form action="{{ route('kandidat.delete', $item->id) }}" method="POST">
                            @csrf
                            @method('DELETE')
                            <div class="flex flex-col">
                                <div class="mb-4 text-black">Apakah Anda yakin ingin menghapus kandidat ini?</div>
                                <button type="submit" class="btn btn-error w-max text-white">Hapus</button>
                            </div>
                        </form>

                    </div>
                </dialog>

                <!-- Detail -->
                <dialog id="show{{ $item->id }}" class="modal">
                    <div class="modal-box bg-white">
                        <h3 class="text-lg font-bold">Hello!</h3>
                        <div class="modal-action">
                            <div class="flex flex-col w-full">
                                <div class="px-4 py-3 w-full bg-gray-300 rounded-lg mb-3">
                                    <label for="" class="text-black font-bold">Nama Kandidat</label>
                                    <p class="text-lg font-medium text-gray-800 mb-2">{{ $item->name_calon }}</p>
                                </div>
                                <div class="px-4 py-3 w-full bg-gray-300 rounded-lg mb-3">
                                    <label for="" class="text-black font-bold">Kelas</label>
                                    <p class="text-lg font-medium text-gray-800 mb-2">{{ $item->kelas_calon }}
                                    </p>
                                </div>
                                <div class="px-4 py-3 w-full bg-gray-300 rounded-lg mb-3">
                                    <label for="" class="text-black font-bold">Visi</label>
                                    <p class="text-lg font-medium text-gray-800 mb-2">{{ $item->visi_calon }}</p>
                                </div>
                                <div class="px-4 py-3 w-full bg-gray-300 rounded-lg mb-3">
                                    <label for="" class="text-black font-bold">Misi</label>
                                    <p class="text-lg font-medium text-gray-800 mb-2">{!! nl2br(e($item->misi_calon)) !!}</p>
                                </div>
                            </div>
                        </div>
                        <button class="btn bg-white text-black hover:bg-red-200"
                            onclick="document.getElementById('show{{ $item->id }}').close()">Close</button>
                    </div>
                </dialog>

                <!-- Edit -->
                <dialog id="edit{{ $item->id }}" class="modal">
                    <div class="modal-box bg-gray-200">
                        <form method="dialog">
                            <button class="btn btn-sm btn-circle btn-ghost absolute right-2 top-2 text-3xl"><i
                                    class="fa-duotone fa-solid fa-circle-xmark text-black"></i></button>
                        </form>
                        <form action="{{ route('edit.kandidat', $item->id) }}" method="POST"
                            enctype="multipart/form-data">
                            @csrf
                            @method('PUT')
                            <div class="mb-2">
                                <label for="formFile" class="mb-2 inline-block  text-black ">Foto Calon</label>
                                <input
                                    class="block w-full text-sm f disabled:pointer-events-none disabled:opacity-60 border boder-gray text-black "
                                    type="file" id="formFile" accept=".jpg,.jpeg,.png" name="image_calon" />
                            </div>
                            <div class="grid grid-cols-2 gap-4 text-white">
                                <div class="mb-4 mx-auto ">
                                <label class="block text-gray-700 font-bold mb-2 text-white" for="name">
                                        Name
                                    </label>
                                    <input
                                        class="shadow appearance-none border rounded w-full py-2 px-3 text-gray-700 leading-tight bg-white "
                                        id="name" type="text" value="{{ $item->name_calon }}" name="name_calon">
                                </div>
                                <div class="mb-4 mx-auto">
                                    <label class="block text-gray-700 font-bold mb-2 text-white" for="kelas">
                                        Kelas
                                    </label>
                                    <input
                                        class="shadow appearance-none border rounded w-full py-2 px-3 text-gray-700 leading-tight bg-white "
                                        id="kelas" type="text" value="{{ $item->kelas_calon }}" name="kelas_calon">
                                </div>
                            </div>
                            <div class="mb-2 ">
                                <label class="block text-gray-700 font-bold mb-2 text-white" for="visi">
                                    Visa Calon
                                </label>
                                <textarea
                                    class="shadow appearance-none border rounded w-full py-2 px-3 text-gray-700 leading-tight bg-white focus:outline-none focus:shadow-outline"
                                    name="visi_calon" id="" cols="30" rows="3">{{ $item->visi_calon }}</textarea>
                            </div>
                            <div class="mb-2">
                                <label class="block text-gray-700 font-bold mb-2 text-white" for="misi">
                                    Misi Calon
                                </label>
                                <textarea
                                    class="shadow appearance-none border rounded w-full py-2 px-3 text-gray-700 leading-tight bg-white focus:outline-none focus:shadow-outline"
                                    name="misi_calon" id="" cols="30" rows="7">{{ $item->misi_calon }}</textarea>
                            </div>
                            <div class="grid grid-cols-2 gap-8">

                                <button type="submit"
                                    class="btn bg-blue-300 hover:bg-transparent  text-white border border-white">Save</button>

                            </div>
                        </form>
                    </div>
                </dialog>

                {{-- @include('dashboard.model-kandidat') --}}
            @endforeach
        </div>

    </div>
    <dialog id="new_kandidat" class="modal">
        <div class="modal-box bg-gray-200">
            <p class="py-4">New Kandidat</p>
            <form action="{{ route('kandidat.create') }}" method="POST" enctype="multipart/form-data">
                @csrf
                <div class="mb-2">
                    <label for="formFile" class="mb-2 inline-block text-neutral-500 ">Foto Calon</label>
                    <input
                        class="mt-2 block w-full text-sm f disabled:pointer-events-none disabled:opacity-60 border boder-gray"
                        type="file" id="formFile" accept=".jpg,.jpeg,.png" name="image_calon" />
                </div>

                <div class="flex flex-row gap-4">
                    <div class="mb-2 w-full">
                        <label class="block text-gray-700 font-bold mb-2" for="name">
                            Nama Calon
                        </label>
                        <input
                            class="bg-white shadow appearance-none border rounded w-full py-2 px-3 text-gray-700 leading-tight focus:outline-none focus:shadow-outline "
                            id="name" type="text" placeholder="your Name Calon" name="name_calon">
                    </div>
                    <div class="mb-2 w-full">
                        <label class="block text-gray-700 font-bold mb-2" for="kelas">
                            Kelas Calon
                        </label>
                        <input
                            class="shadow appearance-none border rounded w-full py-2 px-3 text-gray-700 leading-tight focus:outline-none focus:shadow-outline bg-white "
                            id="kelas" type="text" placeholder="your Kelas Calon" name="kelas_calon">
                    </div>
                </div>
                <div class="mb-2">
                    <label class="block text-gray-700 font-bold mb-2" for="visi">
                        Visa Calon
                    </label>
                    <textarea
                        class="shadow appearance-none border rounded w-full py-2 px-3 text-gray-700 leading-tight focus:outline-none focus:shadow-outline bg-white "
                        name="visi_calon" id="" cols="30" rows="3"></textarea>
                </div>
                <div class="mb-2">
                    <label class="block text-gray-700 font-bold mb-2" for="misi">
                        Misi Calon
                    </label>
                    <textarea
                        class="shadow appearance-none border rounded w-full py-2 px-3 text-gray-700 leading-tight focus:outline-none focus:shadow-outline bg-white "
                        name="misi_calon" id="" cols="30" rows="5"></textarea>
                </div>
                <button type="submit" class="btn btn-primary me-5">Save</button>
                <button class="btn bg-red-200">Close</button>
            </form>
        </div>
    </dialog>
    <script>
        document.getElementById('formFile').addEventListener('change', function() {
            var fileName = this.files[0] ? this.files[0].name : 'Pilih File';
            document.getElementById('fileLabel').textContent = fileName;
        });

        function showEditModal(id) {
            // Hide all other modals
            const modals = document.querySelectorAll('[id^="edit"]');
            modals.forEach(modal => {
                if (modal.id !== 'edit' + id) {
                    modal.classList.add('hidden');
                }
            });

            // Show the selected modal
            document.getElementById('edit' + id).classList.remove('hidden');
        }

        function closeEditModal(id) {
            // Hide the current modal
            document.getElementById('edit' + id).classList.add('hidden');
        }
    </script>
@endsection
