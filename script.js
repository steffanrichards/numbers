function convert() {
    const value = parseFloat(document.getElementById('inputValue').value);
    const inputUnit = document.getElementById('inputUnit').value;
    const outputUnit = document.getElementById('outputUnit').value;
    const density = 1.03; // kg per litre for milk

    if (isNaN(value)) {
        document.getElementById('result').textContent = 'Please enter a valid number.';
        return;
    }

    let result;
    if (inputUnit === outputUnit) {
        result = value;
    } else if (inputUnit === 'litres' && outputUnit === 'kg') {
        result = value * density;
    } else if (inputUnit === 'kg' && outputUnit === 'litres') {
        result = value / density;
    }

    document.getElementById('result').textContent = result.toFixed(2) + ' ' + outputUnit;
}

document.getElementById('convertBtn').addEventListener('click', convert);
