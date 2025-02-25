$(document).ready(function() {
    $('#id_work_type').change(function() {
        let work_type_id = $(this).val();

        if (work_type_id) {
            $.ajax({
                url: ajaxUrls.workTypeUrl, // Define this in your template or script
                data: { 'work_type_id': work_type_id },
                success: function(data) {
                    $('#work_type_description').text(data.description);
                }
            });
        } else {
            $('#work_type_description').text('');
        }
    });
});
