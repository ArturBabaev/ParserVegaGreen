dict_config = {
    'version': 1,
    'disable_existing_loggers': False,
    'formatters': {
        'base': {
            'format': '%(asctime)s | %(levelname)s | %(name)s | %(message)s'
        }
    },
    'handlers': {
        'console': {
            'class': 'logging.StreamHandler',
            'level': 'DEBUG',
            'formatter': 'base'
        },
        'file': {
            'class': 'logging.FileHandler',
            'level': 'ERROR',
            'formatter': 'base',
            'filename': 'logfile.log',
            'mode': 'a'
        }
    },
    'loggers': {
        'parser': {
            'level': 'DEBUG',
            'handlers': ['file', 'console'],
        }
    }
}
