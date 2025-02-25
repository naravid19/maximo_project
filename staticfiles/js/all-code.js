
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
    // ตรวจสอบว่ามีข้อความแจ้งเตือนอยู่แล้วหรือไม่
    if (container.querySelector('.custom-alert')) {
        return; // ถ้ามีอยู่แล้วให้ return ออกไป ไม่ต้องเพิ่มข้อความซ้ำ
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
    // ลบข้อความแจ้งเตือนทั้งหมดในฟอร์ม
    document.querySelectorAll('.custom-alert').forEach(alert => alert.remove());
}
function showUploadLoading(form) {
    const button = document.getElementById('uploadBtn');
    document.body.style.cursor = 'progress';
    button.disabled = true;
    button.classList.add('cursor-not-allowed');
    button.innerHTML = `
        <span class="relative px-5 py-2.5 transition-all ease-in duration-75 bg-white dark:bg-gray-900 rounded-md group-hover:bg-opacity-0">
            <svg aria-hidden="true" role="status" class="inline w-4 h-4 text-gray-200 me-3 animate-spin dark:text-gray-600" viewBox="0 0 100 101" fill="none" xmlns="http://www.w3.org/2000/svg">
                <path d="M100 50.5908C100 78.2051 77.6142 100.591 50 100.591C22.3858 100.591 0 78.2051 0 50.5908C0 22.9766 22.3858 0.59082 50 0.59082C77.6142 0.59082 100 22.9766 100 50.5908ZM9.08144 50.5908C9.08144 73.1895 27.4013 91.5094 50 91.5094C72.5987 91.5094 90.9186 73.1895 90.9186 50.5908C90.9186 27.9921 72.5987 9.67226 50 9.67226C27.4013 9.67226 9.08144 27.9921 9.08144 50.5908Z" fill="currentColor"/>
                <path d="M93.9676 39.0409C96.393 38.4038 97.8624 35.9116 97.0079 33.5539C95.2932 28.8227 92.871 24.3692 89.8167 20.348C85.8452 15.1192 80.8826 10.7238 75.2124 7.41289C69.5422 4.10194 63.2754 1.94025 56.7698 1.05124C51.7666 0.367541 46.6976 0.446843 41.7345 1.27873C39.2613 1.69328 37.813 4.19778 38.4501 6.62326C39.0873 9.04874 41.5694 10.4717 44.0505 10.1071C47.8511 9.54855 51.7191 9.52689 55.5402 10.0491C60.8642 10.7766 65.9928 12.5457 70.6331 15.2552C75.2735 17.9648 79.3347 21.5619 82.5849 25.841C84.9175 28.9121 86.7997 32.2913 88.1811 35.8758C89.083 38.2158 91.5421 39.6781 93.9676 39.0409Z" fill="#1C64F2"/>
            </svg>
            Loading...
        </span>`;
    if (form) {
        form.submit();
    } else {
        console.error('ไม่พบฟอร์ม');
    }
}
document.addEventListener('DOMContentLoaded', function() {
    document.body.style.cursor = 'default';
});

$(document).ready(function() {
    // เมื่อผู้ใช้เลือก Plant Type
    $('#id_plant_type').change(function() {
        let plant_type_id = $(this).val(); // ดึง plant_code ที่เลือก

        // ปิดการใช้งาน dropdown ขณะที่รอข้อมูลจาก server
        $('#id_acttype, #id_site, #id_child_site, #id_work_type, #id_unit').empty().append('<option value="">เลือก</option>');
        $('#id_acttype, #id_site, #id_work_type, #id_unit').prop('disabled', true);
        $('#plant_type_description, #site_description, #child_site_description, #work_type_description, #acttype_description').text('');
        if (plant_type_id) {
            // ส่ง AJAX request ไปที่ server
            $.ajax({
                url: "{% url 'filter_plant_type' %}",
                data: {'plant_type_id': plant_type_id},    // ส่งค่า plant_code ไปยัง backend
                success: function (data) {
                    // ตรวจสอบว่า server ส่งข้อมูลมาถูกต้องหรือไม่
                    if (data) {
                        // อัปเดตคำอธิบายของ plant_type
                        $('#plant_type_description').text(data.plant_type_th); // แสดงคำอธิบาย plant_type
                        
                        // อัปเดต dropdown MNTACT TYPE โดยล้างค่าที่มีอยู่และเพิ่มตัวเลือกใหม่
                        $('#id_acttype').empty().append('<option value="">เลือก</option>');
                        $.each(data.acttypes, function (index, acttype) {
                            $('#id_acttype').append('<option value="' + acttype.id + '">' + acttype.acttype + ' - ' + acttype.description + '</option>');
                        });
                        
                        // อัปเดต dropdown SITE โดยล้างค่าที่มีอยู่และเพิ่มตัวเลือกใหม่
                        $('#id_site').empty().append('<option value="">เลือก</option>');
                        $.each(data.sites, function (index, site) {
                            $('#id_site').append('<option value="' + site.id + '">' + site.site_id + '</option>');
                        });
                        
                        // อัปเดต dropdown WORKTYPE โดยล้างค่าที่มีอยู่และเพิ่มตัวเลือกใหม่
                        $('#id_work_type').empty().append('<option value="">เลือก</option>');
                        $.each(data.work_types, function(index, worktype) {
                            $('#id_work_type').append('<option value="' + worktype.id + '">' + worktype.worktype + ' - ' + worktype.description +'</option>');
                        });

                        // อัปเดต dropdown UNIT
                        $('#id_unit').empty().append('<option value="">เลือก</option>');
                        $.each(data.units, function(index, unit) {
                            $('#id_unit').append('<option value="' + unit.id + '">' + unit.unit_code + '</option>');
                        });

                        // เปิดใช้งาน dropdown หลังจากอัปเดตข้อมูล
                        $('#id_acttype, #id_site, #id_work_type, #id_unit').prop('disabled', false);
                    }
                },
                error: function(xhr, status, error) {
                    $('#id_acttype, #id_site, #id_child_site, #id_work_type, #id_unit').empty().append('<option value="">เลือก</option>');
                    $('#plant_type_description, #site_description, #child_site_description, #work_type_description, #acttype_description').text('');
                    alert("เกิดข้อผิดพลาดในการดึงข้อมูลจาก server: " + error);

                    // เปิดใช้งาน dropdown และตั้งค่าเป็นค่า default
                    $('#id_acttype, #id_site, #id_work_type, #id_unit').prop('disabled', false);
                }
            });
        } else {
            // ถ้าไม่เลือก Plant Type ให้แสดงค่า default
            $('#id_acttype, #id_site, #id_child_site, #id_work_type, #id_unit').empty().append('<option value="">เลือก</option>');
            $('#plant_type_description, #site_description, #child_site_description, #work_type_description, #acttype_description').text('');
            
            // เปิดใช้งาน dropdown และตั้งค่าเป็นค่า default
            $('#id_acttype, #id_site, #id_work_type, #id_unit').prop('disabled', false);
        }
    });
});

$(document).ready(function() {
    // ฟังก์ชันที่ถูกเรียกเมื่อผู้ใช้เลือก Site
    $('#id_site').change(function() {
        let site_id = $(this).val();  // ดึงค่า site_id ที่ผู้ใช้เลือก

        // ปิดการใช้งาน dropdown ของ Child Site ขณะรอข้อมูลจาก server
        $('#id_child_site').prop('disabled', true);
        $('#child_site_description').text('');

        if (site_id) {
            $.ajax({
                url: "{% url 'filter_site' %}",  // URL สำหรับ view ที่ใช้กรอง SITE
                data: {'site_id': site_id},  // ส่ง site_id ไปที่ backend
                dataType: 'json',  // กำหนดประเภทข้อมูลที่คาดว่าจะได้รับเป็น JSON
                success: function (data) {
                    // ตรวจสอบว่ามีข้อมูล ChildSite หรือไม่
                    if (data.child_sites.length > 0) {
                        // ล้างค่าที่มีอยู่ใน dropdown ของ ChildSite และเพิ่มตัวเลือกใหม่
                        $('#id_child_site').empty().append('<option value="">เลือก</option>');
                        $.each(data.child_sites, function(index, child_site) {
                            $('#id_child_site').append('<option value="' + child_site.id + '">' + child_site.site_id + '</option>');
                        });
                    } else {
                        // ถ้าไม่มีข้อมูล ChildSite
                        $('#id_child_site').empty().append('<option value="">ไม่พบข้อมูล Child Site</option>');
                    }

                    // อัปเดตคำอธิบาย site_name
                    $('#site_description').text(data.site_name);  // แสดง site_name ใน HTML
                    // เปิดใช้งาน dropdown ของ Child Site หลังจากอัปเดตข้อมูลเสร็จ
                    $('#id_child_site').prop('disabled', false);
                },
                error: function(xhr, status, error) {
                    // หากเกิดข้อผิดพลาด ให้ล้าง dropdown ของ ChildSite และแสดงข้อความเตือน
                    $('#id_child_site').empty().append('<option value="">เลือก</option>');
                    alert("เกิดข้อผิดพลาดในการดึงข้อมูลจาก server: " + error);
                    // เปิดใช้งาน dropdown ของ Child Site เพื่อให้ผู้ใช้สามารถเลือกใหม่ได้
                    $('#id_child_site').prop('disabled', false);
                }
            });
        } else {
            // ถ้าไม่เลือก Site ให้ล้าง dropdown ของ ChildSite และล้างคำอธิบาย Site
            $('#id_child_site').empty().append('<option value="">เลือก</option>');
            $('#site_description').text('');
            $('#child_site_description').text('');
            $('#id_child_site').prop('disabled', false);  // เปิดใช้งาน dropdown ของ ChildSite
        }
    });
});

$(document).ready(function() {
    // ฟังก์ชันที่ถูกเรียกเมื่อผู้ใช้เลือก Child Site
    $('#id_child_site').change(function() {
        let child_site_id = $(this).val();  // ดึงค่า child_site_id ที่ผู้ใช้เลือก

        if (child_site_id) {
            // ส่ง AJAX request ไปที่ server เพื่อดึงคำอธิบายของ Child Site
            $.ajax({
                url: "{% url 'filter_child_site' %}",  // URL ไปยัง view ที่จะดึงคำอธิบายของ Child Site
                data: {'child_site_id': child_site_id},  // ส่งค่า child_site_id ไปยัง backend
                dataType: 'json',  // กำหนดประเภทข้อมูลที่คาดว่าจะได้รับเป็น JSON
                success: function(data) {
                    // อัปเดตคำอธิบายของ Child Site
                    $('#child_site_description').text(data.description);  // แสดงคำอธิบายใน #child_site_description
                },
                error: function(xhr, status, error) {
                    // หากเกิดข้อผิดพลาด ให้ล้างคำอธิบาย
                    $('#child_site_description').text('ไม่สามารถดึงคำอธิบายได้');
                    alert("เกิดข้อผิดพลาดในการดึงข้อมูลจาก server: " + error);
                }
            });
        } else {
            // ล้างคำอธิบายเมื่อไม่มีการเลือก Child Site
            $('#child_site_description').text('');
        }
    });
});

$(document).ready(function() {

    $('#id_work_type').change(function() {
        let work_type_id = $(this).val();
    
        if(work_type_id){
            $.ajax({
                url: "{% url 'filter_worktype' %}",
                data: {'work_type_id': work_type_id},
                success: function (data){
                    $('#work_type_description').text(data.description);
                }
            });
        } else {
            $('#work_type_description').text('');
        }
    });
});

$(document).ready(function() {
    $('#id_acttype').change(function() {
        let acttype_id = $(this).val();

        if (acttype_id){
            $.ajax({
                url: "{% url 'filter_acttype' %}",
                data: {'acttype_id': acttype_id},
                success: function (data) {
                    $('#acttype_description').text(data.description + ' (' + data.code + ')');
                }
            });
        } else {
            $('#acttype_description').text('');
        }
    });
});

$(document).ready(function() {
    $('#id_wbs').change(function() {
        let wbs_id = $(this).val();
        
        if (wbs_id){
            $.ajax({
                url: "{% url 'filter_wbs' %}",
                data: {'wbs_id': wbs_id},
                success: function (data) {
                    $('#wbs_description').text(data.description);
                }
            });
        } else {
            $('#wbs_description').text('');
        }
    });
});

$(document).ready(function() {
    $('#id_wostatus').change(function() {
        let wostatus_id = $(this).val();
    
        if (wostatus_id) {
            $.ajax({
                url: "{% url 'filter_wostatus' %}",
                data: {'wostatus_id': wostatus_id},
                success: function (data) {
                    $('#wostatus_description').text(data.description);
                }
            });
        } else {
            $('#wostatus_description').text('');
        }
    });
});

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
