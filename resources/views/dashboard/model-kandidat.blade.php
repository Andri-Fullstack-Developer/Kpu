<div class="h-max bg-blue-500 absolute px-7 right-0 -top-48 z-50 hidden py-2 rounded-lg text-white"
    id="edit{{ $item->id }}">
    <div class="w-full flex justify-end">
        <button class="text-lg w-10 h-10 bg-red-300 rounded" onclick="closeEditModal({{ $item->id }})">
            <i class="fa-sharp fa-solid fa-xmark text-white"></i>
        </button>
    </div>
    {{-- <p>id: {{ $item->id }}</p> --}}
    <form action="{{ route('edit.kandidat', $item->id) }}" class="py-10" method="POST" enctype="multipart/form-data">
        @csrf
        @method('PUT')
        <div class="mb-2">
            <label for="formFile" class="mb-2 inline-block  text-black ">Foto Calon</label>
            <input
                class="relative m-0 block w-full min-w-0 flex-auto cursor-pointer rounded border border-solid border-secondary-500 bg-transparent bg-clip-padding px-3 py-[0.32rem] text-base font-normal text-surface transition duration-300 ease-in-out file:-mx-3 file:-my-[0.32rem] file:me-3 file:cursor-pointer file:overflow-hidden file:rounded-none file:border-0 file:border-e file:border-solid file:border-inherit file:bg-transparent file:px-3  file:py-[0.32rem] file:text-surface focus:border-primary focus:text-gray-700 focus:shadow-inset focus:outline-none "
                type="file" id="formFile" accept=".jpg,.jpeg,.png" name="image_calon" />
        </div>
        <div class="grid grid-cols-2 gap-4 text-white">
            <div class="mb-4 mx-auto ">
                <label class="block text-gray-700 font-bold mb-2 text-white" for="name">
                    Name
                </label>
                <input class="shadow appearance-none border rounded w-full py-2 px-3 text-gray-700 leading-tight "
                    id="name" type="text" value="{{ $item->name_calon }}" name="name_calon">
            </div>
            <div class="mb-4 mx-auto">
                <label class="block text-gray-700 font-bold mb-2 text-white" for="kelas">
                    Kelas
                </label>
                <input class="shadow appearance-none border rounded w-full py-2 px-3 text-gray-700 leading-tight "
                    id="kelas" type="text" value="{{ $item->kelas_calon }}" name="kelas_calon">
            </div>
        </div>
        <div class="mb-2 ">
            <label class="block text-gray-700 font-bold mb-2 text-white" for="visi">
                Visa Calon
            </label>
            <textarea
                class="shadow appearance-none border rounded w-full py-2 px-3 text-gray-700 leading-tight focus:outline-none focus:shadow-outline"
                name="visi_calon" id="" cols="30" rows="3">{{ $item->visi_calon }}</textarea>
        </div>
        <div class="mb-2">
            <label class="block text-gray-700 font-bold mb-2 text-white" for="misi">
                Misi Calon
            </label>
            <textarea
                class="shadow appearance-none border rounded w-full py-2 px-3 text-gray-700 leading-tight focus:outline-none focus:shadow-outline"
                name="misi_calon" id="" cols="30" rows="5">{{ $item->misi_calon }}</textarea>
        </div>
        <div class="grid grid-cols-2 gap-8">

            <button type="submit"
                class="btn bg-transparent hover:bg-transparent  text-white border border-white">Save</button>

            {{-- <button onclick="closeEditModal({{ $item->id }})"
                    class="btn btn-outline text-white btn-error">Close</button> --}}
        </div>
    </form>

</div>
