function validateUploadForm(event) {
    event.preventDefault(); // ป้องกันการส่งฟอร์มทันที

    let form = document.getElementById('uploadForm');
    let hasError = false;

    // ลบข้อความแจ้งเตือนทั้งหมดก่อนการตรวจสอบ
    clearAllErrors();

    const yearField = document.getElementById('id_year');
    const frequencyField = document.getElementById('id_frequency');
    const plantTypeField = document.getElementById('id_plant_type');
    const siteField = document.getElementById('id_site');
    const childSiteField = document.getElementById('id_child_site');
    const unitField = document.getElementById('id_unit');
    const wostatusField = document.getElementById('id_wostatus');
    const workTypeField = document.getElementById('id_work_type');
    const acttypeField = document.getElementById('id_acttype');
    const wbsField = document.getElementById('id_wbs');
    const wbsOtherField = document.getElementById('other_wbs');
    const projectidOtherField = document.getElementById('other_projectid');
    const scheduleFileField = document.getElementById('id_schedule_file');
    const locationFileField = document.getElementById('id_location_file');
    const selectedOrderField = document.getElementById('selected_order_input');

    const fieldsToValidate = [
        { field: yearField, message: 'กรุณาเลือกปี ค.ศ.' },
        { field: frequencyField, message: 'กรุณาเลือกความถี่' },
        { field: plantTypeField, message: 'กรุณาเลือกประเภทของโรงไฟฟ้า' },
        { field: siteField, message: 'กรุณาเลือกสังกัดโรงไฟฟ้า' },
        { field: childSiteField, message: 'กรุณาเลือกชื่อโรงไฟฟ้า' },
        { field: unitField, message: 'กรุณาเลือกหมายเลขเครื่องกำเนิดไฟฟ้า' },
        { field: wostatusField, message: 'กรุณาเลือกสถานะใบงาน' },
        { field: workTypeField, message: 'กรุณาเลือกประเภทของใบงาน' },
        { field: acttypeField, message: 'กรุณาเลือกประเภทของกิจกรรมงาน' },
        { field: wbsField, message: 'กรุณาเลือก SUBWBS' },
        { field: scheduleFileField, message: 'กรุณาอัปโหลดไฟล์ Final Schedule' },
        { field: locationFileField, message: 'กรุณาอัปโหลดไฟล์ Location' }
    ];

    // ตรวจสอบฟิลด์ต่างๆ
    fieldsToValidate.forEach(({ field, message }) => {
        if (field.value === '') {
            hasError = true;
            showError(field.closest('.sm\\:col-span-3'), message);
        }
    });

    // ตรวจสอบฟิลด์ wbs และ projectid "other"
    if (wbsField.value === '6') {
        if (wbsOtherField.value === '') {
            hasError = true;
            showError(wbsOtherField.closest('.sm\\:col-span-3'), 'กรุณากรอกข้อมูลในฟิลด์ "Other WBS"');
        }
        if (projectidOtherField.value === '') {
            hasError = true;
            showError(projectidOtherField.closest('.sm\\:col-span-3'), 'กรุณากรอกข้อมูลในฟิลด์ "Other PROJECTID"');
        }
    }

    // ตรวจสอบ selected_order
    if (!selectedOrderField.value.trim()) {
        hasError = true;
        const container = document.getElementById('helper-checkbox-text').parentElement; // ใช้ parentElement แทน closest()
        showError(container, 'กรุณาเลือกตัวเลือก Grouping อย่างน้อยหนึ่งตัวเลือก');
    }

    // ส่งฟอร์มถ้าไม่มีข้อผิดพลาด
    if (!hasError) {
        showUploadLoading(form);
    }
}
function showError(container, message) {
    if (container.querySelector('.custom-alert')) {
        return;
    }

    const alertBox = document.createElement('div');
    alertBox.classList.add('custom-alert', 'flex', 'items-center', 'p-4', 'mt-2', 'text-sm', 'text-red-800', 'border', 'border-red-300', 'rounded-lg', 'bg-red-50');
    alertBox.setAttribute('role', 'alert');

    const messageText = document.createElement('span');
    messageText.textContent = message;
    alertBox.appendChild(messageText);

    container.appendChild(alertBox);
}
function clearAllErrors() {
    document.querySelectorAll('.custom-alert').forEach(alert => alert.remove());
}