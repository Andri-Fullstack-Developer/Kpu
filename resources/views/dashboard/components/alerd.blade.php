@if (session('success'))
    <div id="success-message" class="bg-green-100 border-l-4 border-green-500 text-green-700 p-4 rounded-lg w-full">
        <p class="text-sm font-semibold">Order Status: Confirmed</p>
        <p class="text-xs">{{ session('success') }}</p>
    </div>
@endif
@if (session('error'))
    <div id="error-message" class="bg-red-100 border-l-4 border-red-500 text-red-700 p-4 rounded-lg w-full">
        <p class="text-sm font-semibold">Order Status: Confirmed</p>
        <p class="text-xs">{{ session('error') }}</p>
    </div>
@endif

@if (session('alert'))
    <div id="success-message" class="bg-green-100 border-l-4 border-green-500 text-green-700 p-4 rounded-lg w-full">
        <p class="text-sm font-semibold">Order Status: Confirmed</p>
        <p class="text-xs">{{ session('alert') }}</p>
    </div>
@endif
{{-- @if (session('error'))
    <div id="error-message" class="bg-green-100 border-l-4 border-green-500 text-green-700 p-4 rounded-lg w-full">
        <p class="text-sm font-semibold">Order Status: Confirmed</p>
        <p class="text-xs">{{ session('error') }}</p>
    </div>
@endif --}}

<script>
    setTimeout(function() {
        const successMessage = document.getElementById('success-message');
        const errorMessage = document.getElementById('error-message');

        if (successMessage) {
            successMessage.style.display = 'none';
        }

        if (errorMessage) {
            errorMessage.style.display = 'none';
        }
    }, 5000);
</script>
