# coding: utf8


class InvalidUnitToDXAException(Exception):
    def __init__(self, value):
        self.value = value

    def __str__(self):
        return repr(self.value)
