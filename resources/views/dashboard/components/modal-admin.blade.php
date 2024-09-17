   <!-- Show Delete End Edit -->
   <dialog id="show_Delti{{ $item->id }}" class="modal">
       <div class="modal-box bg-white">
           <form method="dialog">
               <button class="btn btn-sm btn-circle btn-ghost absolute right-2 top-2"><i
                       class="fa-duotone fa-solid fa-circle-xmark text-3xl"></i></button>
           </form>
           <div class="grid grid-cols-2 gap-5 mt-5">

               <button
                   onclick="closeModal('show_Delti{{ $item->id }}'); document.getElementById('EditAdmin{{ $item->id }}').showModal()"
                   class="w-15 h-10 bg-blue-400 rounded-lg text-black "><i
                       class="fa-duotone fa-solid fa-user-pen"></i></button>
               <button
                   onclick="closeModal('show_Delti{{ $item->id }}'); document.getElementById('DeleteAdmin{{ $item->id }}').showModal()"
                   class="w-15 h-10 bg-red-400 rounded-lg text-black"><i
                       class="fa-duotone fa-solid fa-trash"></i></button>
           </div>
       </div>
   </dialog>
   <!-- Edit -->
   <dialog id="EditAdmin{{ $item->id }}" class="modal">
       <div class="modal-box bg-white">
           <form method="dialog">
               <button class="btn btn-sm btn-circle btn-ghost absolute right-2 top-2">✕</button>
           </form>
           <div class="text-start">
               <form action="{{ route('update.admin', $item->id) }}" method="POST">
                   @csrf
                   @method('PUT')
                   <div class="mb-4 ">
                       <label class="block text-gray-700 font-bold mb-2" for="name">
                           Nama Admin
                       </label>
                       <input
                           class="shadow border rounded w-full py-2 px-3 text-gray-700 leading-tight focus:outline-none focus:shadow-outline bg-white"
                           id="name" type="text" placeholder="Your Nama" name="name-admin"
                           value="{{ $item->name_admin }}">
                   </div>
                   <div class="mb-4 ">
                       <label class="block text-gray-700 font-bold mb-2" for="nis-admin">
                           NIS Admin
                       </label>
                       <input
                           class="shadow appearance-none border rounded w-full py-2 px-3 text-gray-700 leading-tight focus:outline-none focus:shadow-outline bg-white"
                           id="nis-admin" type="text" placeholder="Your Nis Admin" name="username-admin"
                           value="{{ $item->username_admin }}">
                   </div>
                   <div class="mb-4 ">
                       <label class="block text-gray-700 font-bold mb-2" for="password-admin">
                           Password Admin
                       </label>
                       <input
                           class="shadow appearance-none border rounded w-full py-2 px-3 text-gray-700 leading-tight focus:outline-none focus:shadow-outline bg-white"
                           id="password-admin" type="text" placeholder="Reset Password Admin" name="password-admin"
                           value="">
                   </div>
                   <button type="submit" class="btn btn-primary text-white">Save</button>
               </form>
           </div>
       </div>
   </dialog>
   <!-- Delete -->
   <dialog id="DeleteAdmin{{ $item->id }}" class="modal">
       <div class="modal-box bg-white">
           <form method="dialog">
               <button class="btn btn-sm btn-circle btn-ghost absolute right-2 top-2">✕</button>
           </form>
           <div class="w-full flex flex-col items-start justify-center ">
               <p class="texx-md mb-4 ">Apakah Anda Yakin Menghapus User?</p>
               <div class="grid grid-cols-2 gap-5">
                   <div>
                       <form action="{{ route('delete.admin', $item->id) }}" method="POST">
                           @csrf
                           @method('DELETE')
                           <button class="px-10 py-4 bg-blue-400 rounded-lg text-white text-bold"
                               type="submit">YA</button>
                       </form>
                   </div>
                   <div>
                       <form action="" method="dialog">
                           <button class="px-10 py-4 bg-red-400 rounded-lg text-white text-bold">Tidak</button>
                       </form>
                   </div>
               </div>
           </div>
       </div>
   </dialog>
