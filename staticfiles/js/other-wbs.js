document.addEventListener('DOMContentLoaded', function () {
    const wbsField = document.getElementById('id_wbs');
    const otherWbsInput = document.getElementById('other_wbs');
    const labelOtherWbs = document.getElementById('label_other_wbs');
    const otherWbsSignal = document.getElementById('other_wbs_signal');
    const otherWbsOuterContainer = document.getElementById('other_wbs_outer_container');

    const otherProjectIdInput = document.getElementById('other_projectid');
    const labelOtherProjectId = document.getElementById('label_other_projectid');
    const otherProjectIdSignal = document.getElementById('other_projectid_signal');
    const otherProjectIdOuterContainer = document.getElementById('other_projectid_outer_container');

    function toggleOtherWbsField() {
        if (wbsField.value === '6') {
            otherWbsInput.disabled = false; // เปิดการใช้งานฟิลด์เมื่อเลือก "other"
            otherWbsInput.placeholder = '';
            otherWbsInput.style.borderBottomColor = '#16a34a';  // เปลี่ยนสีขอบของ input เป็นสีเขียว
            labelOtherWbs.style.color = '#16a34a';  // เปลี่ยนสีข้อความ label เป็นสีเขียว
            otherWbsSignal.classList.remove('hidden');  // แสดงสัญญาณ
            otherWbsOuterContainer.classList.remove('hidden');

            otherProjectIdInput.disabled = false;
            otherProjectIdInput.placeholder = '';
            otherProjectIdInput.style.borderBottomColor = '#16a34a';
            labelOtherProjectId.style.color = '#16a34a';
            otherProjectIdSignal.classList.remove('hidden');
            otherProjectIdOuterContainer.classList.remove('hidden');
        } else {
            otherWbsInput.disabled = true;
            otherWbsInput.placeholder = '';
            otherWbsInput.style.borderBottomColor = '';  // รีเซ็ตสีขอบกลับเป็นค่าเริ่มต้น
            otherWbsInput.value = '';  // ล้างค่าที่กรอก
            labelOtherWbs.style.color = '';  // รีเซ็ตสี label กลับเป็นค่าเริ่มต้น
            otherWbsSignal.classList.add('hidden');  // ซ่อนสัญญาณ
            otherWbsOuterContainer.classList.add('hidden');

            otherProjectIdInput.disabled = true;
            otherProjectIdInput.placeholder = '';
            otherProjectIdInput.style.borderBottomColor = '';
            otherProjectIdInput.value = '';
            labelOtherProjectId.style.color = '';
            otherProjectIdSignal.classList.add('hidden');
            otherProjectIdOuterContainer.classList.add('hidden');
        }
    }

    // ตรวจสอบสถานะเริ่มต้น
    toggleOtherWbsField();

    // ตรวจสอบการเปลี่ยนแปลงในฟิลด์ `wbs`
    wbsField.addEventListener('change', toggleOtherWbsField);
});