$(document).ready(function() {
    $('#id_site').change(function() {
        let site_id = $(this).val();
        $('#id_child_site').prop('disabled', true);
        $('#child_site_description').text('');

        if (site_id) {
            $.ajax({
                url: ajaxUrls.siteFilterUrl, // URL สำหรับดึงข้อมูล Child Site
                data: { 'site_id': site_id },
                dataType: 'json',
                success: function(data) {
                    if (data.child_sites.length > 0) {
                        updateDropdown('#id_child_site', data.child_sites, 'site_id');
                    } else {
                        $('#id_child_site').empty().append('<option value="">ไม่พบข้อมูล Child Site</option>');
                    }
                    $('#site_description').text(data.site_name); // อัปเดตคำอธิบาย Site
                    $('#id_child_site').prop('disabled', false);
                },
                error: function(xhr, status, error) {
                    $('#id_child_site').empty().append('<option value="">เลือก</option>');
                    alert("เกิดข้อผิดพลาด: " + error);
                }
            });
        } else {
            $('#id_child_site').empty().append('<option value="">เลือก</option>');
            $('#site_description').text('');
            $('#child_site_description').text('');
            $('#id_child_site').prop('disabled', false);  // เปิดใช้งาน dropdown ของ ChildSite
        }
    });

    function updateDropdown(selector, data, valueField) {
        $(selector).empty().append('<option value="">เลือก</option>');
        $.each(data, function(index, item) {
            $(selector).append(`<option value="${item.id}">${item[valueField]}</option>`);
        });
    }
});
