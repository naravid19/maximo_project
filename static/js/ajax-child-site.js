$(document).ready(function() {
    $('#id_child_site').change(function() {
        let child_site_id = $(this).val();

        if (child_site_id) {
            $.ajax({
                url: ajaxUrls.childSiteFilterUrl,
                data: { 'child_site_id': child_site_id },
                dataType: 'json',
                success: function(data) {
                    // อัปเดตคำอธิบายของ Child Site
                    $('#child_site_description').text(data.description);
                },
                error: function(xhr, status, error) {
                    // จัดการข้อผิดพลาด
                    console.error("Error fetching child site description:", error);
                    $('#child_site_description').text('ไม่สามารถดึงคำอธิบายได้');
                }
            });
        } else {
            // ล้างคำอธิบายเมื่อไม่มีการเลือก Child Site
            $('#child_site_description').text('');
        }
    });
});
