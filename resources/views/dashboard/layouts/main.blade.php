<!DOCTYPE html>
<html lang="en">

<head>
    <meta charset="UTF-8">
    <meta name="viewport" content="width=device-width, initial-scale=1.0">
    <title>Dashboard Admistrator</title>
    <meta name="description" content="">
    <meta name="csrf-token" content="{{ csrf_token() }}">

    <!-- Daisyui -->
    <link href="{{ asset('assets/vendor/css/dayui.full.min.css') }}" rel="stylesheet" type="text/css">

    <script src="{{ asset('assets/vendor/js/xlsx.full.min.js') }}"></script>

    <!-- Jqury -->
    <script src="{{ asset('assets/vendor/js/jquery-3.6.0.min.js') }}"></script>
    <script src="{{ asset('assets/vendor/js/jquery.dataTables.min.js') }}"></script>


    <!-- Tailwind -->
    <link href="{{ asset('assets/vendor/css/tailwind.min.css') }}" rel="stylesheet">
    <script src="{{ asset('assets/vendor/css/3.4.5.all.min.css') }}"></script>
    <link rel="stylesheet" href="{{ asset('assets/vendor/css/jquery.dataTables.min.css') }}">

    {{-- Icon --}}
    <link rel="stylesheet" href="{{ asset('assets/css/all.min.css') }}">

    {{-- Swalwet --}}
    <link rel="stylesheet" href="https://cdn.jsdelivr.net/npm/sweetalert2@11/dist/sweetalert2.min.css">
    <script src="https://cdn.jsdelivr.net/npm/sweetalert2@11/dist/sweetalert2.all.min.js"></script>

    <style>
        @import url('https://fonts.googleapis.com/css?family=Karla:400,700&display=swap');

        .font-family-karla {
            font-family: karla;
        }

        .bg-sidebar {
            background: #3d68ff;
        }

        .cta-btn {
            color: #3d68ff;
        }

        .upgrade-btn {
            background: #1947ee;
        }

        .upgrade-btn:hover {
            background: #0038fd;
        }

        .active-nav-link {
            background: #1947ee;
        }

        .nav-item:hover {
            background: #1947ee;
        }

        .account-link:hover {
            background: #3d68ff;
        }

        .bg-text {
            background-image: url('{{ asset('assets/images/bgtext.png') }}');
        }

        .dataTables_filter {
            margin-bottom: 13px !important;

        }

        input[type = 'search'] {
            border: 1px solid black !important;
        }

        .dataTables_length select {
            border: 1px solid black !important;
        }
    </style>
</head>

<body class="bg-gray-100 font-family-karla flex">

    <aside class="relative bg-sidebar h-screen w-64 hidden sm:block shadow-xl">
        @include('dashboard/layouts.sidebar')
    </aside>

    <div class="w-full flex flex-col h-screen overflow-y-hidden relative">
        <div class="w-full overflow-x-hidden border-t flex flex-col">
            <div class="sticky top-0 z-50 bg-white shadow-md">
                {{-- @include('dashboard.model-kandidat') --}}
                @include('dashboard/layouts.header')
            </div>
            <main class="w-full flex-grow p-6 bg-text bg-blue-300 overflow-auto">
                <div class="relative min-h-screen">
                    @yield('content')
                </div>
            </main>

        </div>

    </div>

    {{-- Cards --}}
    <script src="{{ asset('assets/vendor/js/chart.umd.min.js') }}"></script>

    <script src="{{ asset('assets/vendor/js/chart.js') }}"></script>

    <!-- AlpineJS -->
    <script src="{{ asset('assets/vendor/js/alpine.min.js') }}" defer></script>
    <!-- Font Awesome -->
    {{-- <script src="https://cdnjs.cloudflare.com/ajax/libs/font-awesome/5.13.0/js/all.min.js"
        integrity="sha256-KzZiKy0DWYsnwMF+X1DvQngQ2/FxF7MF3Ff72XcpuPs=" crossorigin="anonymous"></script> --}}
    <!-- ChartJS -->
    <script src="{{ asset('assets/vendor/js/Chart.min.js') }}"
        integrity="sha256-R4pqcOYV8lt7snxMQO/HSbVCFRPMdrhAFMH+vr9giYI=" crossorigin="anonymous"></script>

</body>

</html>
