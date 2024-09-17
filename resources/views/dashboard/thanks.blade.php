<html>

<head>
    <title>
        Komisi Pemilihan Umum Kabupaten Rembang
    </title>
    <script src="https://cdn.tailwindcss.com"></script>
    <link href="https://cdnjs.cloudflare.com/ajax/libs/font-awesome/5.15.3/css/all.min.css" rel="stylesheet" />
    <link href="https://fonts.googleapis.com/css2?family=Roboto:wght@400;700&amp;display=swap" rel="stylesheet" />
    {{-- Icon --}}
    <link rel="stylesheet" href="{{ asset('assets/css/all.min.css') }}">
    <style>
        body {
            font-family: 'Roboto', sans-serif;
            background-color: #4391D9;
            color: white;
        }

        .banner {
            background-color: #D4AF37;
            padding: 10px 20px;
            border-radius: 5px;
            display: inline-block;
        }

        .header {
            background-color: #4391D9;
            padding: 20px;
            text-align: center;
        }

        .content {
            text-align: center;
            padding: 20px;
        }

        .footer {
            background-color: #4391D9;
            padding: 20px;
            text-align: center;
        }

        .footer img {
            display: inline-block;
            margin: 0 10px;
        }
    </style>
</head>

<body>
    <div class="header">
        <div class="flex justify-between items-center">
            <div class="flex items-center">
                <div class="w-10">
                    <img alt="Logo Komisi Pemilihan Umum Kabupaten Rembang "
                        src="{{ asset('assets/images/kpu.png') }}" />
                </div>
                <div class="ml-4">
                    <h1 class="text-lg font-bold">
                        Komisi Pemilihan Umum
                    </h1>
                    <p class="text-sm">
                        SMPN 1 Sedati
                    </p>
                </div>
            </div>
            <div class="flex items-center w-10">
                <img alt="Pemilu Sarana Integrasi Bangsa" src="{{ asset('assets/images/smpn1.png') }}" />
                {{-- <img alt="14 Februari 2024" src="https://placehold.co/50x50" /> --}}
            </div>
        </div>
    </div>
    <div class="content">
        {{-- <div class="banner">
            <h2 class="text-2xl font-bold">
                Masa Kerja Pantarlih
            </h2>
            <p class="text-lg">
                12 Februari - 11 April 2023
            </p>
        </div> --}}
        <div class="mt-10">
            <h1 class="text-6xl font-bold">
                TERIMA KASIH MEMILIH
            </h1>
        </div>
        <p class="mt-10 text-lg">
            Telah melaksanakan pemilihan kpu Osis. <br> Hasil kerja Pantarlih akan terjaga terawat dan
            bermanfaat. <br> Semoga membawa kemajuan untuk SMPN 1 Sedati
        </p>
    </div>

</body>

</html>
