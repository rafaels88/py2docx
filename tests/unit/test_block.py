# coding: utf-8
#!/usr/bin/env python
import unittest
from py2docx.elements import Block


class BlockTestCase(unittest.TestCase):

    def test_create_block(self):
        b = Block()
        self.assertIsInstance(b, Block)

if __name__ == '__main__':
    unittest.main()
