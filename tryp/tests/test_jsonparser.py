import os
import tempfile
import unittest
from tryp.jsonparser import parse_tryp


class TestParser(unittest.TestCase):
    tryp_file = tempfile.NamedTemporaryFile(delete=False)

    @classmethod
    def tearDownClass(cls):
        os.remove(cls.tryp_file.name)

    def test_parse_tryp(self):
        tryp_file = self.tryp_file
        tryp_file.write('{"crosstab":{}}')
        tryp_file.close()
        json = parse_tryp(tryp_file.name)
        assert isinstance(json, dict)
