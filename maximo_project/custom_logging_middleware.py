import logging

class CustomLoggingMiddleware:
    def __init__(self, get_response):
        self.get_response = get_response

    def __call__(self, request):
        logger = logging.getLogger('django')
        extra = {
            'method': request.method,
            'url': request.build_absolute_uri(),
        }
        response = self.get_response(request)
        extra['status_code'] = response.status_code  # ใส่ค่า status_code หลังจากได้รับ response
        logger = logging.LoggerAdapter(logger, extra)
        return response
