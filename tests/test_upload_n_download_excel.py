# -*- coding: utf-8 -*-

try:
    from urllib import quote
except ImportError:
    from urllib.parse import quote

from testapp import app
import pyexcel as pe
from nose.tools import eq_
from _compact import BytesIO, PY2


_XLSM_MIME = ("application/" +
              "vnd.openxmlformats-officedocument.spreadsheetml.sheet")

FILE_TYPE_MIME_TABLE = {
    "csv": "text/csv",
    "tsv": "text/tab-separated-values",
    "csvz": "application/zip",
    "tsvz": "application/zip",
    "ods": "application/vnd.oasis.opendocument.spreadsheet",
    "xls": "application/vnd.ms-excel",
    "xlsx": _XLSM_MIME,
    "xlsm": "application/vnd.ms-excel.sheet.macroenabled.12"
}


class TestExcelResponse:
    def setUp(self):
        app.config['TESTING'] = True
        self.app = app.test_client()
        self.data = [
            [1, 2, 3],
            [4, 5, 6]
        ]

    def test_download(self):
        for upload_file_type in FILE_TYPE_MIME_TABLE.keys():
            file_name = 'issue15.test.%s' % upload_file_type
            for download_file_type in FILE_TYPE_MIME_TABLE.keys():
                print("Uploading %s Downloading %s" % (
                    upload_file_type, download_file_type))
                io = pe.save_as(dest_file_type=upload_file_type,
                                array=self.data)
                if not PY2:
                    if not isinstance(io, BytesIO):
                        io = BytesIO(io.getvalue().encode('utf-8'))
                response = self.app.post(
                    '/switch/%s' % download_file_type,
                    buffered=True,
                    data={"file": (io, file_name)},
                    content_type="multipart/form-data")
                eq_(response.content_type,
                    FILE_TYPE_MIME_TABLE[download_file_type])
                sheet = pe.get_sheet(file_type=download_file_type,
                                     file_content=response.data)
                sheet.format(int)
                array = sheet.to_array()
                eq_(array, self.data)

    def test_no_file_type(self):
        file_name = 'issue15'
        download_file_type = "xls"
        io = pe.save_as(dest_file_type="xls",
                        array=self.data)
        if not PY2:
            if not isinstance(io, BytesIO):
                io = BytesIO(io.getvalue().encode('utf-8'))
        response = self.app.post(
            '/switch/%s' % download_file_type,
            buffered=True,
            data={"file": (io, file_name)},
            content_type="multipart/form-data")
        eq_(response.content_type, 'text/html')
        eq_(response.status_code, 400)

    def test_override_file_name(self):
        for file_type in FILE_TYPE_MIME_TABLE.keys():
            file_name = 'override_file_name'
            url_encoded_file_name = quote(file_name)
            response = self.app.post('/file_name/%s/%s' % (file_type,
                                                           file_name))
            eq_(response.content_type, FILE_TYPE_MIME_TABLE[file_type])
            eq_(response.headers.get("Content-Disposition", None),
                ("attachment; filename=%s.%s;filename*=utf-8''%s.%s"
                 % (url_encoded_file_name, file_type,
                    url_encoded_file_name, file_type)))

    def test_unicode_file_name(self):
        for file_type in FILE_TYPE_MIME_TABLE.keys():
            file_name = u'中文文件名'
            url_encoded_file_name = quote(file_name.encode('utf-8'))
            response = self.app.post('/file_name/%s/%s' % (file_type,
                                                           file_name))
            eq_(response.content_type, FILE_TYPE_MIME_TABLE[file_type])
            eq_(response.headers.get("Content-Disposition", None),
                ("attachment; filename=%s.%s;filename*=utf-8''%s.%s"
                 % (url_encoded_file_name, file_type,
                    url_encoded_file_name, file_type)))

    def test_utf8_file_name(self):
        for file_type in FILE_TYPE_MIME_TABLE.keys():
            file_name = '中文文件名'
            url_encoded_file_name = quote(file_name)
            response = self.app.post('/file_name/%s/%s' % (file_type,
                                                           file_name))
            eq_(response.content_type, FILE_TYPE_MIME_TABLE[file_type])
            eq_(response.headers.get("Content-Disposition", None),
                ("attachment; filename=%s.%s;filename*=utf-8''%s.%s"
                 % (url_encoded_file_name, file_type,
                    url_encoded_file_name, file_type)))