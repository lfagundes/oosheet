# -*- coding: utf-8 -*-

import sys

def index(name):
    letters = [ l for l in name.upper() ]
    letters.reverse()
    index = 0
    power = 0
    for letter in letters:
        index += (1 + ord(letter) - ord('A')) * pow(ord('Z') - ord('A') + 1, power)
        power += 1
    return index - 1

def name(index):
    name = []
    letters = [ chr(ord('A')+i) for i in range(26) ]

    while index > 0:
        i = index % 26
        index = int(index/26) - 1
        name.append(letters[i])

    if index == 0:
        name.append('A')

    name.reverse()
    return ''.join(name)            

class Wrapper(object):
    def __init__(self, wrapped):
        self.wrapped = wrapped
    def __getattr__(self, name):
        try:
            return getattr(self.wrapped, name)
        except AttributeError:
            if not name == name.upper():
                raise ImportError
            return name

sys.modules[__name__] = Wrapper(sys.modules[__name__])
            
