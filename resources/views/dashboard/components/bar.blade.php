<canvas id="kandidatBarChart"></canvas>

<script>
    $(document).ready(function() {
        // Data dari server (diambil dari PHP dan di-encode ke JSON)
        var kandidatData = @json($kandidatData);
        var pemilihPerKelas = @json($pemilihPerKelas);

        // Konfigurasi Bar Chart Jumlah Suara per Kandidat
        var labelsKandidat = kandidatData.map(function(kandidat) {
            return kandidat.name;
        });

        var dataVotesKandidat = kandidatData.map(function(kandidat) {
            return kandidat.votes;
        });

        const dataKandidat = {
            labels: labelsKandidat,
            datasets: [{
                label: 'Jumlah Suara',
                data: dataVotesKandidat,
                backgroundColor: [
                    'rgba(54, 162, 235, 0.2)',
                    'rgba(255, 99, 132, 0.2)',
                    'rgba(75, 192, 192, 0.2)',
                    'rgba(153, 102, 255, 0.2)',
                    'rgba(255, 159, 64, 0.2)',
                    // Tambahkan lebih banyak warna jika diperlukan
                ],
                borderColor: [
                    'rgba(54, 162, 235, 1)',
                    'rgba(255, 99, 132, 1)',
                    'rgba(75, 192, 192, 1)',
                    'rgba(153, 102, 255, 1)',
                    'rgba(255, 159, 64, 1)',
                    // Tambahkan lebih banyak warna jika diperlukan
                ],
                borderWidth: 1
            }]
        };

        const configKandidat = {
            type: 'bar',
            data: dataKandidat,
            options: {
                responsive: true,
                plugins: {
                    legend: {
                        position: 'top',
                    },
                    title: {
                        display: true,
                        text: 'Jumlah Suara per Kandidat'
                    }
                }
            },
        };

        // Render bar chart untuk kandidat
        var kandidatBarChart = new Chart(
            document.getElementById('kandidatBarChart'),
            configKandidat
        );

        // Konfigurasi Bar Chart Jumlah Pemilih per Kelas
        var labelsKelas = pemilihPerKelas.map(function(item) {
            return item.kelas;
        });

        var dataPemilihKelas = pemilihPerKelas.map(function(item) {
            return item.jumlah_pemilih;
        });

        const dataKelas = {
            labels: labelsKelas,
            datasets: [{
                label: 'Jumlah Pemilih',
                data: dataPemilihKelas,
                backgroundColor: 'rgba(255, 206, 86, 0.2)',
                borderColor: 'rgba(255, 206, 86, 1)',
                borderWidth: 1
            }]
        };

        const configKelas = {
            type: 'bar',
            data: dataKelas,
            options: {
                responsive: true,
                plugins: {
                    legend: {
                        position: 'top',
                    },
                    title: {
                        display: true,
                        text: 'Jumlah Pemilih per Kelas'
                    }
                }
            },
        };

        // Render bar chart untuk kelas
        var kelasBarChart = new Chart(
            document.getElementById('kelasBarChart'),
            configKelas
        );
    });
</script>
