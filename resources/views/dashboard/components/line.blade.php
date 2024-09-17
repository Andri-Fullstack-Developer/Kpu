<canvas id="voteChart"></canvas>


<script>
    $(document).ready(function() {
        var ctx = $('#voteChart')[0].getContext('2d');

        // Ambil data JSON dari Laravel
        var minutes = {!! json_encode($votesData->pluck('minute')) !!};
        var totalVotes = {!! json_encode($votesData->pluck('total_votes')) !!};

        console.log(minutes); // Untuk memastikan array minutes
        console.log(totalVotes); // Untuk memastikan array totalVotes

        // Render chart dengan Chart.js
        var chart = new Chart(ctx, {
            type: 'line',
            data: {
                labels: minutes, // Label pada sumbu X (menit)
                datasets: [{
                    label: 'Jumlah Vote per Menit',
                    backgroundColor: 'rgba(75, 192, 192, 0.2)',
                    borderColor: 'rgba(75, 192, 192, 1)',
                    data: totalVotes, // Data jumlah vote
                }]
            },
            options: {
                responsive: true,
                scales: {
                    x: {
                        title: {
                            display: true,
                            text: 'Waktu (HH:mm)'    
                        }
                    },
                    y: {
                        title: {
                            display: true,
                            text: 'Jumlah Vote'
                        }
                    }
                }
            }
        });
    });
</script>
