# -*- coding: utf-8 -*-

def index(name):
    letters = [l for l in name.upper()]
    letters.reverse()
    index = 0
    power = 0
    for letter in letters:
        index += (1+ord(letter)-ord('A')) * pow(ord('Z')-ord('A')+1, power)
        power += 1
    return index - 1


def name(index):
    name = []
    letters = [chr(ord('A')+i) for i in range(26)]

    while index > 0:
        i = index % 26
        index = int(index/26) - 1
        name.append(letters[i])

    if index == 0:
        name.append('A')

    name.reverse()
    return ''.join(name)
