# Copyright (C) Michael Cook <michael@waxrat.com>
import pytest
from pytest import approx
import os
import subprocess

@pytest.fixture(scope='module')
def example_py():
    code = subprocess.check_output(['./xlsx2py', 'example.xlsx'], encoding='utf-8')
    with open('example.py', 'wt') as fh:
        fh.write(code)
    yield
    os.remove('example.py')

def test_rectangular_range(example_py):
    from example import Sheet_Sheet1 as sh1
    assert sh1.A1() == 'rectangular range'
    assert sh1.A12() == approx(-1281.6) # min
    assert sh1.A13() == approx(1612.8)  # max
    assert sh1.A14() == approx(64.17)   # average
    assert sh1.A15() == approx(3208.5)  # sum

def test_range_with_empty_cells(example_py):
    from example import Sheet_Sheet1 as sh1
    assert sh1.A17() == 'range with empty cells'
    assert sh1.A19() == approx(12)      # min
    assert sh1.A20() == approx(56)      # max
    assert sh1.A21() == approx(34)      # average
    assert sh1.A22() == approx(102)     # sum

def test_formula_references_another_sheet(example_py):
    from example import Sheet_Sheet1 as sh1
    assert sh1.B24() == 'formula references another sheet'
    assert sh1.A24() == approx(356)

def test_rept(example_py):
    from example import Sheet_Sheet2 as sh2
    assert sh2.A1() == approx(10)
    assert sh2.A2() == approx(13.5)
    assert sh2.A3() == approx(17)
    assert sh2.A4() == approx(20.5)
    assert sh2.A5() == approx(24)
    assert sh2.A6() == approx(27.5)
    assert sh2.A7() == approx(31)
    assert sh2.A8() == approx(34.5)
    assert sh2.A9() == approx(178)

    assert sh2.B1() == '||||||||||'
    assert sh2.B2() == '|||||||||||||'
    assert sh2.B3() == '|||||||||||||||||'
    assert sh2.B4() == '||||||||||||||||||||'
    assert sh2.B5() == '||||||||||||||||||||||||'
    assert sh2.B6() == '|||||||||||||||||||||||||||'
    assert sh2.B7() == '|||||||||||||||||||||||||||||||'
    assert sh2.B8() == '||||||||||||||||||||||||||||||||||'
    assert sh2.B9() == 'total'

def test_literals(example_py):
    from example import Sheet_Foo_Bar as fb
    assert fb.A1() == 10
    assert fb.B1() == 'start'
    assert fb.A2() == 3.5
    assert fb.B2() == 'increment'

def test_date(example_py):
    from example import Sheet_Foo_Bar as fb
    assert fb.A4() == approx(44138)              # =DATE(2020, 11, 3)
    assert fb.A6() == approx(44138.5383449074)   # =DATE(2020, 11, 3, 12, 55)

def test_time(example_py):
    from example import Sheet_Foo_Bar as fb
    assert fb.A5() == approx(0.538352575862974)  # =TIME(12, 55, 13, 662555)

def test_if(example_py):
    from example import Sheet_Sheet1 as sh1

    assert sh1.A27()
    assert sh1.B27()
    assert sh1.C27()
    assert sh1.D27()
    assert sh1.E27()
    assert sh1.F27()

    assert not sh1.A28()
    assert not sh1.B28()
    assert not sh1.C28()
    assert not sh1.D28()
    assert not sh1.E28()
    assert not sh1.F28()

    assert sh1.A29() == "yup"
    assert sh1.B29() == "yup"
    assert sh1.C29() == "yup"
    assert sh1.D29() == "yup"
    assert sh1.E29() == "yup"
    assert sh1.F29() == "yup"

    assert sh1.A30() == "yup"
    assert sh1.B30() == "yup"
    assert sh1.C30() == "yup"
    assert sh1.D30() == "yup"
    assert sh1.E30() == "yup"
    assert sh1.F30() == "yup"

    assert sh1.A31() == "yup"
    assert sh1.B31() == "yup"
    assert sh1.C31() == "yup"
    assert sh1.D31() == "yup"
    assert sh1.E31() == "yup"
    assert sh1.F31() == "yup"

    assert not sh1.A32()
    assert not sh1.B32()
    assert not sh1.C32()
    assert not sh1.D32()
    assert not sh1.E32()
    assert not sh1.F32()

    assert sh1.A34()        # =TRUE()
    assert not sh1.A35()    # =FALSE()

def test_isformula(example_py):
    from example import Sheet_Sheet1 as sh1

    assert sh1.B37() == 'isformula – formula cell'
    assert sh1.A37()
    assert sh1.B38() == 'isformula – number cell'
    assert not sh1.A38()
    assert sh1.B39() == 'isformula – text cell'
    assert not sh1.A38()
    assert sh1.B40() == 'isformula – empty cell'
    assert not sh1.A38()
