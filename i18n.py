"""
Translate sheets and field names.

This class handle the translation from english to hebrew.
"""

import json
import system


class i18n(object):
    """
    i18n handler.

    Translating the fields and sheets names.
    """

    fields = {}
    sheets = {}

    def __init__(self):
        """Constructing."""
        self.fields = {}
        self.sheets = {}

    def set_fields(self, fields):
        """
        Set the fields.

        When you need to set different values for the fields translation you
        can use this method.

        :param fields:
            The new fields to set.
        """
        pass

    def translate_field(self, field):
        """
        Translate a given field form english to hebrew.

        :param field:
            The field name for translation.
        """
        pass

    def set_sheets(self, sheets):
        """
        Set the sheets.

        When you need to set different values for the sheets translation you
        can use this method.

        :param fields:
            The new sheets to set.
        """
        self.sheets = sheets

    def translate_sheet(self, sheet):
        """
        Translate a given sheet form english to hebrew.

        :param field:
            The field name for translation.
        """
        pass
