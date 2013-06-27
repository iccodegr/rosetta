FORMAT_FILTERS = {
    'pdf': {
        'filters': {
            '.doc': 'writer_pdf_Export',
            '.docx': 'writer_pdf_Export',
            '.txt': 'writer_pdf_Export',
        },
        'extension': 'pdf'
    }
}


def get_filter(source, fmt):
    if not fmt in FORMAT_FILTERS:
        return None

    fmt = FORMAT_FILTERS[fmt]
    return fmt['filters'][source], fmt['extension']
