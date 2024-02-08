from functions import compile_regex
import re
import pytest

@pytest.fixture
def patterns():
    # List of search patterns for testing the function
    return ["abc", "123", "XYZ"]

@pytest.fixture
def compiled_patterns(patterns):
    # Calling the function to compile the patterns
    return compile_regex(patterns)

def test_compile_regex_length(compiled_patterns, patterns):
    # Verify that the number of compiled patterns is the same as the number of input patterns
    assert len(compiled_patterns) == len(patterns)

def test_compile_regex_type(compiled_patterns):
    # Verify that each element in the list of compiled patterns is indeed a compiled regular expression
    for compiled_pattern in compiled_patterns:
        assert isinstance(compiled_pattern, re.Pattern)
