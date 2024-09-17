<canvas id="groupedBarChart"></canvas>


<script>
    $(document).ready(function() {
        // Data dari server (diambil dari PHP dan di-encode ke JSON)
        var groupedBarData = @json($groupedBarData);

        // Debug data dari server (opsional, untuk memastikan data benar)
        console.log(groupedBarData);

        var labelsGrouped = groupedBarData.map(function(item) {
            return item.kelas;
        });

        var datasetsGrouped = [];
        groupedBarData[0].candidates.forEach(function(candidate, index) {
            datasetsGrouped.push({
                label: candidate.name_calon,
                data: groupedBarData.map(function(item) {
                    return item.candidates[index].votes;
                }),
                backgroundColor: 'rgba(' + (index * 50 + 100) + ', ' + (index * 50 + 50) +
                    ', ' + (index * 50 + 150) + ', 0.7)'
            });
        });

        const dataGrouped = {
            labels: labelsGrouped,
            datasets: datasetsGrouped
        };

        const configGrouped = {
            type: 'bar',
            data: dataGrouped,
            options: {
                responsive: true,
                plugins: {
                    legend: {
                        position: 'top',
                    },
                    title: {
                        display: true,
                        text: 'Perbandingan Suara per Kelas'
                    }
                },
                scales: {
                    x: {
                        stacked: true,
                    },
                    y: {
                        stacked: true
                    }
                }
            },
        };

        // Render grouped bar chart
        var groupedBarChart = new Chart(
            document.getElementById('groupedBarChart'),
            configGrouped
        );
    });
</script>
