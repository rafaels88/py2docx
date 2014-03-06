# coding: utf8


class InvalidUnitToDXAException(Exception):
    def __init__(self, value):
        self.value = value

    def __str__(self):
        return repr(self.value)


class Unit(object):

    @classmethod
    def to_dxa(cls, val):
        if len(val) < 2:
            return 0
        else:
            unit = val[-2:]
            val = int(val.rstrip(unit))
            if val == 0:
                return 0

            if unit == 'cm':
                return cls.cm_to_dxa(val)
            elif unit == 'in':
                return cls.in_to_dxa(val)
            elif unit == 'pt':
                return cls.pt_to_dxa(val)
            else:
                raise InvalidUnitToDXAException("Unit to DXA should be " +
                                                "Centimeters(cm), " +
                                                "Inches(in) or " +
                                                "Points(pt)")

    @classmethod
    def pixel_to_emu(cls, pixel):
        return int(round(pixel * 12700))

    @classmethod
    def cm_to_dxa(cls, centimeters):
        inches = centimeters / 2.54
        points = inches * 72
        dxa = points * 20
        return dxa

    @classmethod
    def in_to_dxa(cls, inches):
        points = inches * 72
        dxa = cls.pt_to_dxa(points)
        return dxa

    @classmethod
    def pt_to_dxa(cls, points):
        dxa = points * 20
        return dxa
