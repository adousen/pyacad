#!/usr/bin/env python
# -*- coding: utf-8 -*-
"""
    pyautocad.types
    ~~~~~~~~~~~~~~~

    3D Points and and other AutoCAD data types.

    :copyright: (c) 2012 by Roman Haritonov.
    :license: BSD, see LICENSE.txt for more details.
"""
import win32com.client
import pythoncom
import array
import operator
import math

class AVariantArray(win32com.client.VARIANT):
    'Create a array of variant that applys to AUTOCAD via com'
    def __init__(self, x_or_seq):
        if isinstance(x_or_seq, (array.array, list, tuple)):
            self.varianttype = pythoncom.VT_ARRAY | pythoncom.VT_VARIANT
            self._value = x_or_seq        
        else:
            self.varianttype = pythoncom.VT_ARRAY | pythoncom.VT_VARIANT
            self._value = [x_or_seq]
    def __repr__(self):
        return "win32com.client.VARIANT(%r, %r)" % (self.varianttype, self._value)
    def __str__(self):
        return "AVariantArray(%r)" % (self._value)


class AIntArray(win32com.client.VARIANT):
    def __init__(self, x_or_seq):
        if isinstance(x_or_seq, (array.array, list, tuple)):
            self.varianttype = pythoncom.VT_ARRAY | pythoncom.VT_I2
            elf._value = x_or_seq        
        else:
            self.varianttype = pythoncom.VT_ARRAY | pythoncom.VT_I2
            self._value = [x_or_seq]
    def __repr__(self):
        return "win32com.client.VARIANT(%r, %r)" % (self.varianttype, self._value)
    def __str__(self):
        return "AIntArray(%r)" % (self._value)
    
class ALongArray(win32com.client.VARIANT):
    def __init__(self, x_or_seq):
        if isinstance(x_or_seq, (array.array, list, tuple)):
            self.varianttype = pythoncom.VT_ARRAY | pythoncom.VT_I4
            elf._value = x_or_seq        
        else:
            self.varianttype = pythoncom.VT_ARRAY | pythoncom.VT_I4
            self._value = [x_or_seq]
    def __repr__(self):
        return "win32com.client.VARIANT(%r, %r)" % (self.varianttype, self._value)
    def __str__(self):
        return "ALongArray(%r)" % (self._value)
    
class AShortArray(win32com.client.VARIANT):
    def __init__(self, x_or_seq):
        if isinstance(x_or_seq, (array.array, list, tuple)):
            self.varianttype = pythoncom.VT_ARRAY | pythoncom.VT_R4
            elf._value = x_or_seq        
        else:
            self.varianttype = pythoncom.VT_ARRAY | pythoncom.VT_R4
            self._value = [x_or_seq]
    def __repr__(self):
        return "win32com.client.VARIANT(%r, %r)" % (self.varianttype, self._value)
    def __str__(self):
        return "AShortArray(%r)" % (self._value)
    
class ADoubleArray(win32com.client.VARIANT):
    def __init__(self, x_or_seq):
        if isinstance(x_or_seq, (array.array, list, tuple)):
            self.varianttype = pythoncom.VT_ARRAY | pythoncom.VT_R8
            elf._value = x_or_seq        
        else:
            self.varianttype = pythoncom.VT_ARRAY | pythoncom.VT_R8
            self._value = [x_or_seq]
    def __repr__(self):
        return "win32com.client.VARIANT(%r, %r)" % (self.varianttype, self._value)
    def __str__(self):
        return "ADoubleArray(%r)" % (self._value)
    
class AStringArray(win32com.client.VARIANT):
    def __init__(self, x_or_seq):
        if isinstance(x_or_seq, (array.array, list, tuple)):
            self.varianttype = pythoncom.VT_ARRAY | pythoncom.VT_BSTR
            elf._value = x_or_seq        
        else:
            self.varianttype = pythoncom.VT_ARRAY | pythoncom.VT_BSTR
            self._value = [x_or_seq]
    def __repr__(self):
        return "win32com.client.VARIANT(%r, %r)" % (self.varianttype, self._value)
    def __str__(self):
        return "AStringArray(%r)" % (self._value)
    
class APoint(win32com.client.VARIANT):
    """ 3D point with basic geometric operations and support for passing as a
        parameter for `AutoCAD` Automation functions

    Usage::

        >>> p1 = APoint(10, 10)
        >>> p2 = APoint(20, 20)
        >>> p1 + p2
        APoint(30.00, 30.00, 0.00)

    Also it supports iterable as parameter::

        >>> APoint([10, 20, 30])
        APoint(10.00, 20.00, 30.00)
        >>> APoint(range(3))
        APoint(0.00, 1.00, 2.00)

    Supported math operations: `+`, `-`, `*`, `/`, `+=`, `-=`, `*=`, `/=`::

        >>> p = APoint(10, 10)
        >>> p + p
        APoint(20.00, 20.00, 0.00)
        >>> p + 10
        APoint(20.00, 20.00, 10.00)
        >>> p * 2
        APoint(20.00, 20.00, 0.00)
        >>> p -= 1
        >>> p
        APoint(9.00, 9.00, -1.00)

    It can be converted to `tuple` or `list`::

        >>> tuple(APoint(1, 1, 1))
        (1.0, 1.0, 1.0)

    """
    def __init__(self, x_or_seq, y=0.0, z=0.0):
        if isinstance(x_or_seq, (array.array, list, tuple)) and len(x_or_seq) == 3:
            self.varianttype = pythoncom.VT_ARRAY | pythoncom.VT_R8
            self._value = x_or_seq        
        else:
            self.varianttype = pythoncom.VT_ARRAY | pythoncom.VT_R8
            self._value = (x_or_seq, y, z)

    @property
    def x(self):
        """ x coordinate of 3D point"""
        return self._value[0]

    @x.setter
    def x(self, value):
        self._value[0] = value

    @property
    def y(self):
        """ y coordinate of 3D point"""
        return self._value[1]

    @y.setter
    def y(self, value):
        self._value[1] = value

    @property
    def z(self):
        """ z coordinate of 3D point"""
        return self._value[2]

    @z.setter
    def z(self, value):
        self._value[2] = value

##    def distance_to(self, other):
##        """ Returns distance to `other` point
##
##        :param other: :class:`APoint` instance or any sequence of 3 coordinates
##        """
##        return distance(self, other)
##
##    def __add__(self, other):
##        return self.__left_op(self, other, operator.add)
##
##    def __sub__(self, other):
##        return self.__left_op(self, other, operator.sub)
##
##    def __mul__(self, other):
##        return self.__left_op(self, other, operator.mul)
##
##    def __div__(self, other):
##        return self.__left_op(self, other, operator.div)
##
##    __radd__ = __add__
##    __rsub__ = __sub__
##    __rmul__ = __mul__
##    __rdiv__ = __div__
##
##    def __neg__(self):
##        return self.__left_op(self, -1, operator.mul)
##
##    def __left_op(self, p1, p2, op):
##        if isinstance(p2, (float, int)):
##            return APoint(op(p1[0], p2), op(p1[1], p2), op(p1[2], p2))
##        return APoint(op(p1[0], p2[0]), op(p1[1], p2[1]), op(p1[2], p2[2]))
##
##    def __iadd__(self, p2):
##        return self.__iop(p2, operator.add)
##
##    def __isub__(self, p2):
##        return self.__iop(p2, operator.sub)
##
##    def __imul__(self, p2):
##        return self.__iop(p2, operator.mul)
##
##    def __idiv__(self, p2):
##        return self.__iop(p2, operator.div)
##
##    def __iop(self, p2, op):
##        if isinstance(p2, (float, int)):
##            self[0] = op(self[0], p2)
##            self[1] = op(self[1], p2)
##            self[2] = op(self[2], p2)
##        else:
##            self[0] = op(self[0], p2[0])
##            self[1] = op(self[1], p2[1])
##            self[2] = op(self[2], p2[2])
##        return self
##
    def __repr__(self):
        return "win32com.client.VARIANT(%r, %r)" % (self.varianttype, self._value)
##
    def __str__(self):
        return 'APoint(%.2f, %.2f, %.2f)' % (self.x,self.y,self.z)
##
##    def __eq__(self, other):
##        if not isinstance(other, (array.array, list, tuple)):
##            return False
##        return tuple(self) == tuple(other)


##def distance(p1, p2):
##    """ Returns distance between two points `p1` and `p2`
##    """
##    return math.sqrt((p1[0] - p2[0]) ** 2 +
##                     (p1[1] - p2[1]) ** 2 +
##                     (p1[2] - p2[2]) ** 2)
##
##
### next functions can accept parameters as aDouble(1, 2, 3)
### or as list or tuple aDouble([1, 2, 3])
##def aDouble(*seq):
##    """ Returns :class:`array.array` of doubles ('d' code) for passing to AutoCAD
##
##    For 3D points use :class:`APoint` instead.
##    """
##    return _sequence_to_comtypes('d', *seq)
##
##
##def aInt(*seq):
##    """ Returns :class:`array.array` of ints ('l' code) for passing to AutoCAD
##    """
##    return _sequence_to_comtypes('l', *seq)
##
##
##def aShort(*seq):
##    """ Returns :class:`array.array` of shorts ('h' code) for passing to AutoCAD
##    """
##    return _sequence_to_comtypes('h', *seq)
##
##
##def _sequence_to_comtypes(typecode='d', *sequence):
##    if len(sequence) == 1:
##        return array.array(typecode, sequence[0])
##    return array.array(typecode, sequence)
