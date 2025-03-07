$(document).ready(function() {
    $('#id_wostatus').change(function() {
        let wostatus_id = $(this).val();

        if (wostatus_id) {
            $.ajax({
                url: ajaxUrls.wostatusFilterUrl, // Define this in your template or script
                data: { 'wostatus_id': wostatus_id },
                success: function(data) {
                    $('#wostatus_description').text(data.description);
                }
            });
        } else {
            $('#wostatus_description').text('');
        }
    });
});
