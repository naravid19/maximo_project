$(document).ready(function() {
    $('#id_plant_type').change(function() {
        let plant_type_id = $(this).val();

        $('#id_acttype, #id_site, #id_child_site, #id_work_type, #id_unit').empty().append('<option value="">เลือก</option>');
        $('#id_acttype, #id_site, #id_work_type, #id_unit').prop('disabled', true);
        $('#plant_type_description, #site_description, #child_site_description, #work_type_description, #acttype_description').text('');

        if (plant_type_id) {
            $.ajax({
                url: ajaxUrls.plantTypeUrl, // Define this in your template or script
                data: { 'plant_type_id': plant_type_id },
                success: function(data) {
                    if (data) {
                        $('#plant_type_description').text(data.plant_type_th);
                        updateDropdown('#id_acttype', data.acttypes, 'acttype', 'description');
                        updateDropdown('#id_site', data.sites, 'site_id');
                        updateDropdown('#id_work_type', data.work_types, 'worktype', 'description');
                        updateDropdown('#id_unit', data.units, 'unit_code');
                        $('#id_acttype, #id_site, #id_work_type, #id_unit').prop('disabled', false);
                    }
                },
                error: function(xhr, status, error) {
                    handleError('#id_acttype, #id_site, #id_work_type, #id_unit');
                    alert("เกิดข้อผิดพลาด: " + error);
                }
            });
        } else {
            resetDropdowns();
        }
    });

    function updateDropdown(selector, data, valueField, textField) {
        $(selector).empty().append('<option value="">เลือก</option>');
        $.each(data, function(index, item) {
            $(selector).append(`<option value="${item.id}">${item[valueField]}${textField ? ' - ' + item[textField] : ''}</option>`);
        });
    }

    function handleError(selectors) {
        $(selectors).empty().append('<option value="">เลือก</option>');
        $('#plant_type_description, #site_description, #child_site_description, #work_type_description, #acttype_description').text('');
    }

    function resetDropdowns() {
        $('#id_acttype, #id_site, #id_child_site, #id_work_type, #id_unit').empty().append('<option value="">เลือก</option>');
        $('#id_acttype, #id_site, #id_work_type, #id_unit').prop('disabled', false);
        $('#plant_type_description, #site_description, #child_site_description, #work_type_description, #acttype_description').text('');
    }
});