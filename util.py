# coding: utf8


class Unit(object):

    @staticmethod
    def pixel_to_dxa(pixels):
        inches = pixels / 75
        points = inches * 72
        dxa = points * 20
        return dxa
