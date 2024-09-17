<!DOCTYPE html>
<html lang="en">

<head>
    <meta charset="UTF-8">
    <meta name="viewport" content="width=device-width, initial-scale=1.0">
    <title>Login KPU Students</title>
    <link href="https://cdn.jsdelivr.net/npm/daisyui@4.12.10/dist/full.min.css" rel="stylesheet" type="text/css">
    <script src="https://cdn.tailwindcss.com"></script>
    <!-- SweetAlert2 CDN -->
    <link rel="stylesheet" href="https://cdn.jsdelivr.net/npm/sweetalert2@11/dist/sweetalert2.min.css">

</head>

<body>
    <div
        class="w-screen min-h-screen flex items-center justify-center bg-[#4391D9] dark:bg-gray-800 px-4 sm:px-6 lg:px-8 relative">
        <div class="absolute md:top-18 md:right-0 md:left-[20%] top-24 -button-10 -left-0 z-50">
            <img class="w-3/5 md:w-1/4" src="{{ asset('assets/images/ayo_memilih.png') }}" alt="">
        </div>
        <div class="py-3 sm:max-auto md:w-3/12 -z-8">
            <div class="min-h-96 px-8 py-6 mt-4 text-left bg-white dark:bg-gray-900  rounded-xl shadow-lg">
                <!-- Display error messages -->

                <form action="{{ route('login.submit') }}" method="POST">
                    @csrf
                    <div class="flex flex-col justify-center items-center h-full select-none">
                        <div class="flex flex-col items-center justify-c    enter gap-2 mb-8">
                            <a href="https://smpn1sedati.sch.id/" target="_blank">
                                <img src="{{ asset('assets/images/kpu.png') }}" alt="Logo KPu" class="w-8">
                            </a>
                            <p class="m-0 text-[16px] font-semibold dark:text-white">Login KPU Students</p>
                            <span class="m-0 text-xs max-w-[90%] text-center text-[#8B8E98]">
                                KPU OSIS Election of OSIS SMPN 1 Sedati.
                            </span>
                        </div>
                        <div class="w-full flex flex-col gap-2">
                            <label class="font-semibold text-xs text-gray-400 ">Username</label>
                            <input
                                class="border rounded-lg px-3 py-2 mb-5 text-sm w-full outline-none dark:border-gray-500 dark:bg-gray-900"
                                placeholder="Username" name="username">
                        </div>
                    </div>
                    <div class="w-full flex flex-col gap-2">
                        <label class="font-semibold text-xs text-gray-400 ">Password</label>
                        <input type="password"
                            class="border rounded-lg px-3 py-2 mb-5 text-sm w-full outline-none dark:border-gray-500 dark:bg-gray-900"
                            placeholder="••••••••" name="password">
                    </div>
                    <div className="mt-5">
                        <button type="submit"
                            class="py-1 px-8 bg-blue-500 hover:bg-blue-800 focus:ring-offset-blue-200 text-white w-full transition ease-in duration-200 text-center text-base font-semibold shadow-md focus:outline-none focus:ring-2 focus:ring-offset-2 rounded-lg cursor-pointer select-none">Login</button>
                    </div>
                </form>
            </div>
        </div>
    </div>
    <!-- SweetAlert2 JavaScript -->
    <script src="https://cdn.jsdelivr.net/npm/sweetalert2@11/dist/sweetalert2.all.min.js"></script>
    @if (session('alert'))
        <script>
            Swal.fire({
                icon: 'error',
                title: 'Oops...',
                text: '{{ session('alert') }}',
            });
        </script>
    @endif
</body>

</html>
