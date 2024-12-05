// Array เพื่อเก็บลำดับการเลือก
const selectedOrder = [];

function toggleArrangeOptions(checkbox) {
    const noArrangeCheckbox = document.getElementById('no-arrange-checkbox');
    const arrangeOptions = document.getElementById('arrange-options');

    if (checkbox.checked) {
        arrangeOptions.classList.remove('hidden');
        noArrangeCheckbox.checked = false;
        const index = selectedOrder.indexOf('no_arrange');
        if (index > -1) selectedOrder.splice(index, 1);
    } else {
        arrangeOptions.classList.add('hidden');
        document.querySelectorAll('#arrange-options input[type="checkbox"]').forEach(childCheckbox => {
            childCheckbox.checked = false;
            const index = selectedOrder.indexOf(childCheckbox.value);
            if (index > -1) selectedOrder.splice(index, 1);
        });
    }
    updateSelectedOrder();
}

function handleCheckboxSelection(checkbox) {
    if (checkbox.checked) {
        if (!selectedOrder.includes(checkbox.value)) {
            selectedOrder.push(checkbox.value);
        }

        if (checkbox.value === 'no_arrange') {
            document.getElementById('arrange-group-checkbox').checked = false;
            document.getElementById('arrange-options').classList.add('hidden');
            document.querySelectorAll('#arrange-options input[type="checkbox"]').forEach(childCheckbox => {
                if (childCheckbox.checked) {
                    childCheckbox.checked = false;
                    const index = selectedOrder.indexOf(childCheckbox.value);
                    if (index > -1) selectedOrder.splice(index, 1);
                }
            });
        }
    } else {
        const index = selectedOrder.indexOf(checkbox.value);
        if (index > -1) selectedOrder.splice(index, 1);
    }
    updateSelectedOrder();
}

function updateSelectedOrder() {
    document.getElementById('selected_order_input').value = selectedOrder.join(',');
}