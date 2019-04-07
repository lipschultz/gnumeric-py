"""
Gnumeric-py: Reading and writing gnumeric files with python
Copyright (C) 2017 Michael Lipschultz

This program is free software: you can redistribute it and/or modify
it under the terms of the GNU General Public License as published by
the Free Software Foundation, either version 3 of the License, or
(at your option) any later version.

This program is distributed in the hope that it will be useful,
but WITHOUT ANY WARRANTY; without even the implied warranty of
MERCHANTABILITY or FITNESS FOR A PARTICULAR PURPOSE. See the
GNU General Public License for more details.

You should have received a copy of the GNU General Public License
along with this program. If not, see <http://www.gnu.org/licenses/>.
"""

import unittest

from gnumeric.expression_evaluation import evaluate


class EvaluationTests(unittest.TestCase):
    ANY_CELL = None

    def test_it_can_start_with_a_plus_sign(self):
        actual = evaluate('+-2*3', self.ANY_CELL)
        self.assertEqual(-6, actual)

    def test_it_evaluates_integers(self):
        actual = evaluate('=54', self.ANY_CELL)
        self.assertEqual(54, actual)
        self.assertIsInstance(actual, int)

    def test_it_evaluates_floats(self):
        actual = evaluate('=5.4', self.ANY_CELL)
        self.assertEqual(5.4, actual)

    def test_it_evaluates_text(self):
        actual = evaluate('="test"', self.ANY_CELL)
        self.assertEqual('test', actual)

    def test_it_evaluates_true(self):
        actual = evaluate('=TRUE', self.ANY_CELL)
        self.assertEqual(True, actual)

    def test_it_evaluates_false(self):
        actual = evaluate('=FALSE', self.ANY_CELL)
        self.assertEqual(False, actual)

    def test_basic_arithmetic_evaluation(self):
        cases = (
            ('=-2', -2),
            ('=2+3', 5),
            ('=2-3', -1),
            ('=2*3', 6),
            ('=2/3', 2/3),
            ('=2^3', 2**3),
            ('=-2+3*4-10/5', 8),
            ('=2*(8-3)^2^3', 2*(8-3)**2**3),
        )
        for case, expected_result in cases:
            actual = evaluate(case, self.ANY_CELL)
            self.assertEqual(expected_result, actual, f'Result mismatch on {case}')

    def test_numeric_logical_evaluation(self):
        cases = (
            ('=2<3', True),
            ('=2<=3', True),
            ('=2>3', False),
            ('=2>=3', False),
            ('=2=3', False),
            ('=2<>3', True),
            ('=2<1+5', True),
        )
        for case, expected_result in cases:
            actual = evaluate(case, self.ANY_CELL)
            self.assertEqual(expected_result, actual, f'Result mismatch on {case}')

    def test_string_logical_evaluation(self):
        cases = (
            ('="case"<"test"', True),
            ('="case"<="test"', True),
            ('="case">"test"', False),
            ('="case">="test"', False),
            ('="case"="test"', False),
            ('="case"<>"test"', True),
        )
        for case, expected_result in cases:
            actual = evaluate(case, self.ANY_CELL)
            self.assertEqual(expected_result, actual, f'Result mismatch on {case}')

    def test_string_logical_evaluation_is_case_insensitive(self):
        cases = (
            ('="case"<"CASE"', False),
            ('="case"<="CASE"', True),
            ('="case">"CASE"', False),
            ('="case">="CASE"', True),
            ('="case"="CASE"', True),
            ('="case"<>"CASE"', False),
        )
        for case, expected_result in cases:
            actual = evaluate(case, self.ANY_CELL)
            self.assertEqual(expected_result, actual, f'Result mismatch on {case}')

    def test_numbers_always_less_than_text_and_bools(self):
        cases = (
            ('=2="2"', False),
            ('=2<>"2"', True),
            ('=50<"2"', True),
            ('=50<="2"', True),
            ('=2>"1"', False),
            ('=2>="1"', False),

            (f'=2=TRUE', False),
            (f'=2<>TRUE', True),
            (f'=50<TRUE', True),
            (f'=50<=TRUE', True),
            (f'=2>TRUE', False),
            (f'=2>=TRUE', False),

            (f'=2=FALSE', False),
            (f'=2<>FALSE', True),
            (f'=50<FALSE', True),
            (f'=50<=FALSE', True),
            (f'=2>FALSE', False),
            (f'=2>=FALSE', False),
        )
        for case, expected_result in cases:
            actual = evaluate(case, self.ANY_CELL)
            self.assertEqual(expected_result, actual, f'Result mismatch on {case}')

    def test_text_always_less_than_bools(self):
        cases = (
            (f'="cat"=TRUE', False),
            (f'="cat"<>TRUE', True),
            (f'="cat"<TRUE', True),
            (f'="cat"<=TRUE', True),
            (f'="cat">TRUE', False),
            (f'="cat">=TRUE', False),

            (f'="cat"=FALSE', False),
            (f'="cat"<>FALSE', True),
            (f'="cat"<FALSE', True),
            (f'="cat"<=FALSE', True),
            (f'="cat">FALSE', False),
            (f'="cat">=FALSE', False),

            (f'=""=FALSE', False),
            (f'=""<>FALSE', True),
            (f'=""<FALSE', True),
            (f'=""<=FALSE', True),
            (f'="">FALSE', False),
            (f'="">=FALSE', False),
        )
        for case, expected_result in cases:
            actual = evaluate(case, self.ANY_CELL)
            self.assertEqual(expected_result, actual, f'Result mismatch on {case}')

    def test_boolean_logical_evaluation(self):
        cases = (
            (f'=TRUE=TRUE', True),
            (f'=TRUE<>TRUE', False),
            (f'=TRUE<TRUE', False),
            (f'=TRUE<=TRUE', True),
            (f'=TRUE>TRUE', False),
            (f'=TRUE>=TRUE', True),

            (f'=TRUE=FALSE', False),
            (f'=TRUE<>FALSE', True),
            (f'=TRUE<FALSE', False),
            (f'=TRUE<=FALSE', False),
            (f'=TRUE>FALSE', True),
            (f'=TRUE>=FALSE', True),

            (f'=FALSE=FALSE', True),
            (f'=FALSE<>FALSE', False),
            (f'=FALSE<FALSE', False),
            (f'=FALSE<=FALSE', True),
            (f'=FALSE>FALSE', False),
            (f'=FALSE>=FALSE', True),
        )
        for case, expected_result in cases:
            actual = evaluate(case, self.ANY_CELL)
            self.assertEqual(expected_result, actual, f'Result mismatch on {case}')

    def test_text_concatenation(self):
        cases = (
            ('="cat"&"dog"', 'catdog'),
            ('=2&"cat"', '2cat'),
            ('="cat"&2', 'cat2'),
            ('="cat"&-2+3*4-10/5', 'cat8'),
            ('=(2<3)&"cat"', 'TRUEcat'),
            ('=(2>3)&"cat"', 'FALSEcat'),
        )
        for case, expected_result in cases:
            actual = evaluate(case, self.ANY_CELL)
            self.assertEqual(expected_result, actual, f'Result mismatch on {case}')
