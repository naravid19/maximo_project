$(document).ready(function() {
    $('#id_acttype').change(function() {
        let acttype_id = $(this).val();

        if (acttype_id) {
            $.ajax({
                url: ajaxUrls.actTypeUrl, // Define this in your template or script
                data: { 'acttype_id': acttype_id },
                success: function(data) {
                    $('#acttype_description').text(`${data.description} (${data.code})`);
                }
            });
        } else {
            $('#acttype_description').text('');
        }
    });
});
