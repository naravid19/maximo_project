$(document).ready(function() {
    $('#id_wbs').change(function() {
        let wbs_id = $(this).val();

        if (wbs_id) {
            $.ajax({
                url: ajaxUrls.wbsFilterUrl, // Define this in your template or script
                data: { 'wbs_id': wbs_id },
                success: function(data) {
                    $('#wbs_description').text(data.description);
                }
            });
        } else {
            $('#wbs_description').text('');
        }
    });
});
