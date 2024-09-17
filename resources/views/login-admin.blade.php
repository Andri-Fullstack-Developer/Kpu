<!DOCTYPE html>
<html lang="en">

<head>
    <meta charset="UTF-8">
    <meta name="viewport" content="width=device-width, initial-scale=1.0">
    <title>Login</title>
    <link href="https://cdn.jsdelivr.net/npm/daisyui@4.12.10/dist/full.min.css" rel="stylesheet" type="text/css">
    <script src="https://cdn.tailwindcss.com"></script>
</head>

<body>
    <div
        class="w-screen min-h-screen flex items-center justify-center bg-[#4391D9] dark:bg-gray-800 px-4 sm:px-6 lg:px-8">
        <div class="relative py-3 sm:max-auto md:w-3/12">
            <div class="min-h-96 px-8 py-6 mt-4 text-left bg-white dark:bg-gray-900  rounded-xl shadow-lg">
                <div class="z-999">
                    <!-- Display error messages -->
                    @if ($errors->any())
                        <div class="alert alert-danger">
                            <ul>
                                @foreach ($errors->all() as $error)
                                    <li>{{ $error }}</li>
                                @endforeach
                            </ul>
                        </div>
                    @endif

                    <!-- Display MySQL error message -->
                    @if (session('mysql_error'))
                        <div class="alert alert-danger">
                            {{ session('mysql_error') }}
                        </div>
                    @endif
                </div>
                <form action="{{ route('admin.submit') }}" method="POST">
                    @csrf
                    <div class="flex flex-col justify-center items-center h-full select-none">
                        <div class="flex flex-col items-center justify-center gap-2 mb-8">
                            <a href="https://smpn1sedati.sch.id/" target="_blank">
                                <img src="{{ asset('assets/images/kpu.png') }}" alt="Logo" class="w-8">
                            </a>
                            <p class="m-0 text-[16px] font-semibold dark:text-white">Login KPU Teachers</p>
                            <span class="m-0 text-xs max-w-[90%] text-center text-[#8B8E98]">
                                KPU OSIS Election of OSIS SMPN 1 Sedati.
                            </span>
                        </div>
                        <div class="w-full flex flex-col gap-2">
                            <label class="font-semibold text-xs text-gray-400 ">Username</label>
                            <input
                                class="border rounded-lg px-3 py-2 mb-5 text-sm w-full outline-none dark:border-gray-500 dark:bg-gray-900"
                                placeholder="Username" name="username_admin">
                        </div>
                    </div>
                    <div class="w-full flex flex-col gap-2">
                        <label class="font-semibold text-xs text-gray-400 ">Password</label>
                        <input type="password"
                            class="border rounded-lg px-3 py-2 mb-5 text-sm w-full outline-none dark:border-gray-500 dark:bg-gray-900"
                            placeholder="••••••••" name="password_admin">
                    </div>
                    <div className="mt-5">
                        <button type="submit"
                            class="py-1 px-8 bg-blue-500 hover:bg-blue-800 focus:ring-offset-blue-200 text-white w-full transition ease-in duration-200 text-center text-base font-semibold shadow-md focus:outline-none focus:ring-2 focus:ring-offset-2 rounded-lg cursor-pointer select-none">Login</button>
                    </div>
                </form>
            </div>
        </div>
    </div>
    </div>
</body>

</html>
