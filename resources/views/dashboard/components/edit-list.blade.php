<div class="modal-box bg-gray-300">
    <button type="button" class="btn btn-sm btn-circle btn-ghost absolute right-2 top-2"
        onclick="document.getElementById('editUser{{ $user->id }}').close()">
        <i class="fa-duotone fa-solid fa-circle-xmark text-xl text-black"></i>
    </button>
    <div class="text-start">
        <form action="{{ route('editUser.create', $user->id) }}" method="POST">
            @csrf
            @method('PUT')
            <div class="mb-4 ">
                <label class="block text-gray-700 font-bold mb-2" for="name">
                    Nama Kandidat
                </label>
                <input
                    class="shadow appearance-none border rounded w-full py-2 px-3 text-gray-700 leading-tight focus:outline-none focus:shadow-outline bg-white"
                    id="name" type="text" placeholder="Nama Kandidat" name="name"
                    value="{{ $user->name }}">
            </div>
            <div class="mb-4 ">
                <label class="block text-gray-700 font-bold mb-2" for="kelas">
                    Kelas Kandidat
                </label>
                <input
                    class="shadow appearance-none border rounded w-full py-2 px-3 text-gray-700 leading-tight focus:outline-none focus:shadow-outline bg-white"
                    id="kelas" type="text" placeholder="Kelas Kandidat" name="kelas"
                    value="{{ $user->kelas }}">
            </div>
            <button type="submit" class="btn btn-primary text-white">Save</button>
        </form>
    </div>
</div>
