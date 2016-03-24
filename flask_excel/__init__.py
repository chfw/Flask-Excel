"""
    flask_excel
    ~~~~~~~~~~~~~~~~~~~

    A flask extension that provides one application programming interface
    to read and write data in different excel file formats

    :copyright: (c) 2015 by Onni Software Ltd.
    :license: New BSD License
"""
from flask import Flask, Request, Response
import pyexcel as pe
import pyexcel_webio as webio


class ExcelRequest(webio.ExcelInputInMultiDict, Request):
    """
    Mix in pyexcel's webio function signatures to Flask request
    """
    def get_file_tuple(self, field_name):
        # will upload multi files with the same filename
        for filehandle in self.files.getlist(field_name):
            filename = filehandle.filename
            extension = filename.split(".")[1]
            yield extension, filehandle

    def get_array(self, **keywords):
        """
        Get a list of lists from the file

        :param sheet_name: For an excel book, there could be multiple
                           sheets. If it is left unspecified, the
                           sheet at index 0 is loaded. For 'csv',
                           'tsv' file, *sheet_name* should be None anyway.
        :param keywords: additional key words
        :returns: A list of lists
        """
        result = []
        for params in self.get_params(**keywords):
            result.extend(pe.get_array(**params))
        return result

    def get_params(self, field_name=None, **keywords):
        """
        Load the single sheet from named form field
        """
        for file_type, file_handle in self.get_file_tuple(field_name):
            if file_type is not None and file_handle is not None:
                file_content = file_handle.read()
                file_content = file_content.decode("gbk").encode("utf-8")
                keywords = {
                    'file_type': file_type,
                    'file_content': file_content
                }
                yield keywords
            else:
                raise Exception("Invalid parameters")


# Plug-in the custom request to Flask
Flask.request_class = ExcelRequest


def _make_response(content, content_type, status, file_name=None):
    """
    Custom response function that is called by pyexcel-webio
    """
    response = Response(content, content_type=content_type, status=status)
    if file_name:
        response.headers["Content-Disposition"] = "attachment; filename=%s" % (file_name)
    return response


webio.ExcelResponse = _make_response


from pyexcel_webio import (
    make_response,
    make_response_from_array,
    make_response_from_dict,
    make_response_from_records,
    make_response_from_book_dict,
    make_response_from_a_table,
    make_response_from_query_sets,
    make_response_from_tables
)
